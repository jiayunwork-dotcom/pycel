# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import collections
import hashlib
import itertools as it
import json
import logging
import math
import os
import pickle
from numbers import Number

import networkx as nx
import numpy as np
from ruamel.yaml import YAML

from pycel.excelformula import ExcelFormula
from pycel.excelutil import (
    AddressCell,
    AddressRange,
    ERROR_CODES,
    flatten,
    is_address,
    iterative_eval_tracker,
    list_like,
    MAX_COL,
    MAX_ROW,
    SPILL_ERROR,
)
from pycel.excelwrapper import ExcelOpxWrapper, ExcelOpxWrapperNoData

REF_START = '=_REF_("'
REF_END = '")'
REF_FORMAT = REF_START + '{}' + REF_END

Mismatch = collections.namedtuple('Mismatch', 'original calced formula')

pycel_logger = logging.getLogger('pycel')


class ExcelCompiler:
    """Class responsible for taking an Excel spreadsheet and compiling it
    to an instance that can be serialized to disk, and executed
    independently of excel.
    """

    save_file_extensions = ('pkl', 'pickle', 'yml', 'yaml', 'json')

    def __init__(self, filename=None, excel=None, plugins=None, cycles=None):
        """ Build a compiler instance to organize the formula for a workbook

        :param filename: Excel filename to load from (xlsx or `to_file`)
        :param excel: Opened instance of ExcelWrapper or openpyxl workbook
        :param plugins: module paths for plugin lib functions
        :param cycles: Override workbook iterative calculation settings
        """

        self._eval = None

        if excel:
            # if we are running as an excel addin, this gets passed to us
            if not isinstance(excel, (ExcelOpxWrapper, _CompiledImporter)):
                excel = ExcelOpxWrapperNoData(excel)
            self.excel = excel
            self.filename = excel.filename
            self.hash = None
        else:
            self.excel = ExcelOpxWrapper(filename=filename)
            self.excel.load()
            self.filename = filename

        # grab a copy of the current hash
        self._excel_file_md5_digest = self._compute_excel_file_md5_digest

        self.log = pycel_logger

        # directed graph for cell dependencies
        self.dep_graph = nx.DiGraph()

        # cell address to Cell mapping, cells and ranges already built
        self.cell_map = {}

        # cells, ranges and graph_edges that need to be built
        self.graph_todos = []
        self.range_todos = []

        self.extra_data = None
        self.conditional_formats = {}
        self._formula_cells_dict = {}
        self._plugin_modules = plugins

        # Dynamic array spill bookkeeping (e.g. SEQUENCE)
        # spill_map:   anchor address -> addresses the anchor spilled to
        # spill_owners: spilled address -> owning anchor address
        # both maps only ever contain plain address strings
        self.spill_map = {}
        self.spill_owners = {}
        # anchors currently being resolved, to guard against recursion
        self._spill_resolving = set()
        # cached dynamic array anchor addresses per workbook sheet
        self._dynamic_anchor_cache = {}

        # Setup to be able to evaluate circular references
        self.cycles = cycles
        if hasattr(self.excel, 'workbook'):
            if self.excel.workbook.calculation is None:
                wb_cycles = False
            else:
                wb_cycles = bool(self.excel.workbook.calculation.iterate)
            if self.cycles is None:
                self.cycles = wb_cycles
            elif wb_cycles != bool(self.cycles):
                self.log.warning(
                    f"Initialized with cycles: {self.cycles}, while workbook says: {wb_cycles}")

            if self.cycles:
                self.cycles = dict(
                    iterations=self.excel.workbook.calculation.iterateCount,
                    tolerance=self.excel.workbook.calculation.iterateDelta,
                )

        elif self.cycles:
            assert isinstance(self.cycles, dict)
            assert self.cycles.keys() == {'iterations', 'tolerance'}

        self.Cell = _CycleCell if self.cycles else _Cell
        self.evaluate = (self._evaluate_iterative if self.cycles else
                         self._evaluate_non_iterative)

    def __getstate__(self):
        # code objects are not serializable
        state = dict(self.__dict__)
        to_removes = '_eval excel log graph_todos range_todos ' \
                     'conditional_formats'.split()
        for to_remove in to_removes:
            if to_remove in state:    # pragma: no branch
                state[to_remove] = None
        return state

    def __setstate__(self, d):
        self.__dict__.update(d)
        self.log = pycel_logger

    @staticmethod
    def _compute_file_md5_digest(filename):
        if not os.path.exists(filename):
            return None
        else:
            hash_md5 = hashlib.md5()
            with open(filename, "rb") as f:
                for chunk in iter(lambda: f.read(4096), b""):
                    hash_md5.update(chunk)
            return hash_md5.hexdigest()

    @property
    def _compute_excel_file_md5_digest(self):
        return self._compute_file_md5_digest(self.filename)

    @property
    def hash_matches(self):
        current_hash = self._compute_excel_file_md5_digest
        return self._excel_file_md5_digest == current_hash

    @property
    def eval(self):
        if self._eval is None:
            eval_ctx = ExcelFormula.build_eval_context(
                self._evaluate, self._evaluate_range,
                self.log, plugins=self._plugin_modules)

            if self.cycles:
                def _eval(cell, cse_array_address=None):
                    cell.start_calcs()
                    return eval_ctx(
                        cell.formula, cse_array_address=cse_array_address)

            else:
                def _eval(cell, cse_array_address=None):
                    return eval_ctx(
                        cell.formula, cse_array_address=cse_array_address)

            self._eval = _eval

        return self._eval

    @classmethod
    def _filename_has_extension(cls, filename):
        return next((extension for extension in cls.save_file_extensions
                     if filename.endswith(extension)), None)

    def _to_text(self, filename=None, is_json=False):
        """Serialize to a json/yaml file"""
        extra_data = {} if self.extra_data is None else self.extra_data

        def cell_value(a_cell):
            if a_cell.formula and a_cell.formula.python_code:
                return '=' + a_cell.formula.python_code
            elif isinstance(a_cell.value, np.float64):
                return float(a_cell.value)
            else:
                return a_cell.value

        extra_data.update(dict(
            cycles=self.cycles,
            excel_hash=self._excel_file_md5_digest,
            cell_map=dict(sorted(
                ((addr, cell_value(cell))
                 for addr, cell in self.cell_map.items() if cell.serialize),
                key=lambda x: AddressRange(x[0]).sort_key
            )),
            # dynamic array spill ownership, spilled values ride along
            # in the cell_map above as plain cell values
            spill_map={
                anchor: list(spilled)
                for anchor, spilled in self.spill_map.items()
            },
            # serialize the workbook filename (not the serialization path)
            filename=self.filename,
        ))
        if not filename:
            filename = self.filename + ('.json' if is_json else '.yml')

        # hash the current file to see if this function makes any changes
        existing_hash = (self._compute_file_md5_digest(filename)
                         if os.path.exists(filename) else None)

        if not is_json:
            with open(filename, 'w') as f:
                ymlo = YAML()
                ymlo.width = 120
                ymlo.dump(extra_data, f)
        else:
            with open(filename, 'w') as f:
                json.dump(extra_data, f, indent=4)

        del extra_data['cell_map']

        # hash the newfile, return True if it changed, this is only reliable
        # on pythons which have ordered dict (CPython 3.6 & python 3.7+)
        return (existing_hash is None or
                existing_hash != self._compute_file_md5_digest(filename))

    @classmethod
    def _from_text(cls, filename, is_json=False):
        """deserialize from a json/yaml file"""

        if not is_json:
            if not filename.split('.')[-1].startswith('y'):
                filename += '.yml'
        else:
            if not filename.endswith('.json'):  # pragma: no branch
                filename += '.json'

        with open(filename, 'r') as f:
            data = YAML().load(f)

        excel = _CompiledImporter(filename, data)
        excel_compiler = cls(excel=excel, cycles=data.pop('cycles', False))
        excel.compiler = excel_compiler

        def add_line_numbers(cell_addr, line_number):
            formula = excel_compiler.cell_map[cell_addr].formula
            if formula is not None:
                formula.lineno = line_number
                formula.filename = filename

        # populate the cells
        range_todos = []
        for address, python_code in data['cell_map'].items():
            lineno = data['cell_map'].lc.data[address][0] + 1
            address = AddressRange(address)
            if address.is_range:
                range_todos.append((address, lineno))
            else:
                excel_compiler._make_cells(address)
                add_line_numbers(address.address, lineno)

        # populate the ranges and dependant graph
        for address, lineno in range_todos:
            excel_compiler._make_cells(address)
            add_line_numbers(address.address, lineno)

        excel_compiler._process_gen_graph()
        del data['cell_map']

        # restore dynamic array spill ownership
        excel_compiler._restore_spills(data.pop('spill_map', None))

        # process the rest of the data from the file
        excel_compiler._excel_file_md5_digest = data['excel_hash']
        del data['excel_hash']
        excel_compiler.extra_data = data

        # remove "excel" file references for GC
        excel_compiler.excel = None
        return excel_compiler

    def to_file(self, filename=None, file_types=('pkl', 'yml')):
        """ Save the spreadsheet to a file so it can be loaded later w/o excel

        :param filename: filename to save as, defaults to xlsx_name + file_type
        :param file_types: one or more of: pkl, pickle, yml, yaml, json

        If the filename has one of the expected extensions, then this
        parameter is ignored.

        The text file formats (yaml and json) provide the benefits of:
            1. Can `diff` subsequent version of xlsx to monitor changes.
            2. Can "debug" the generated code.
                Since the compiled code is marked with the line number in
                the text file and will be shown by debuggers and stack traces.
            3. The file size on disk is somewhat smaller than pickle files

        The pickle file format provides the benefits of:
            1. Much faster to load (5x to 10x)
            2. ...  (no #2, speed is the thing)
        """

        filename = filename or self.filename
        extension = self._filename_has_extension(filename)
        if extension:
            file_types = (extension, )
        elif isinstance(file_types, str):
            file_types = (file_types, )

        unknown_types = tuple(ft for ft in file_types
                              if ft not in self.save_file_extensions)
        if unknown_types:
            raise ValueError(f"Unknown file types: {' '.join(unknown_types)}")

        pickle_extension = next((ft for ft in file_types
                                 if ft.startswith('p')), None)
        non_pickle_extension = next((ft for ft in file_types
                                     if not ft.startswith('p')), None)
        extra_extensions = tuple(ft for ft in file_types if ft not in (
            pickle_extension, non_pickle_extension))

        if extra_extensions:
            raise ValueError(
                'Only allowed one pickle extension and one text extension. '
                f'Extras: {extra_extensions}')

        is_json = non_pickle_extension and non_pickle_extension[0] == 'j'

        # round trip through yaml/json to strip out junk
        text_name = filename
        if not text_name.endswith(non_pickle_extension or '.yml'):
            text_name += '.' + (non_pickle_extension or 'yml')
        text_changed = self._to_text(text_name, is_json=is_json)

        # save pickle file if requested and has changed
        if pickle_extension:
            if not filename.endswith(pickle_extension):
                filename += '.' + pickle_extension

            if text_changed or not os.path.exists(filename):
                excel_compiler = self._from_text(text_name, is_json=is_json)
                if non_pickle_extension not in file_types:
                    os.unlink(text_name)

                with open(filename, 'wb') as f:
                    pickle.dump(excel_compiler, f)

    @classmethod
    def from_file(cls, filename, plugins=None):
        """ Load the spreadsheet saved by `to_file`

        :param filename: filename to load from, can be xlsx_name
        :param plugins: module paths for plugin lib functions
        """

        extension = cls._filename_has_extension(filename) or next(
            (ext for ext in cls.save_file_extensions
             if os.path.exists(filename + '.' + ext)), None)

        if not extension:
            raise ValueError(f"Unrecognized file type or compiled file not found: '{filename}'")

        if not filename.endswith(extension):
            filename += '.' + extension

        if extension[0] == 'p':
            with open(filename, 'rb') as f:
                excel_compiler = pickle.load(f)
        else:
            excel_compiler = cls._from_text(
                filename, is_json=extension == 'json')

        excel_compiler.excel = _CompiledImporter('', {
            'filename': excel_compiler.filename,
            'cell_map': excel_compiler.cell_map,
        })
        excel_compiler.range_todos = []
        excel_compiler.graph_todos = []

        excel_compiler._plugin_modules = plugins
        return excel_compiler

    def export_to_dot(self, filename=None):
        try:
            # test pydot is importable  (optionally installed)
            import pydot  # noqa: F401
        except ImportError:
            raise ImportError("Package 'pydot' is not installed")

        from networkx.drawing.nx_pydot import write_dot
        filename = filename or (self.filename + '.dot')
        write_dot(self.dep_graph, filename)

    def export_to_gexf(self, filename=None):
        # networkx<2.7 still references np.float_, removed in numpy 2.x
        import numpy as np
        if not hasattr(np, 'float_'):  # pragma: no cover
            np.float_ = np.float64

        from networkx.readwrite.gexf import write_gexf
        filename = filename or (self.filename + '.gexf')
        write_gexf(self.dep_graph, filename)

    def plot_graph(self, layout_type='spring_layout'):
        try:
            # test matplotlib is importable  (optionally installed)
            import matplotlib.pyplot as plt
        except ImportError:
            raise ImportError("Package 'matplotlib' is not installed")

        pos = getattr(nx, layout_type)(self.dep_graph, iterations=2000)
        nx.draw_networkx_nodes(self.dep_graph, pos)
        nx.draw_networkx_edges(self.dep_graph, pos, arrows=True)
        nx.draw_networkx_labels(self.dep_graph, pos)
        plt.show()

    def set_value(self, address, value, set_as_range=False):
        """ Set the value of one or more cells or ranges

        :param address: `str`, `AddressRange`, `AddressCell` or a tuple, list
            or an iterable of these three
        :param value: value to set.  This can be a value or a tuple/list
            which matches the shapes needed for the given address/addresses
        :param set_as_range: With a single range address and a list like value,
            set to true to set the entire range to the inserted list.
        """

        if list_like(value) and not set_as_range:
            value = tuple(flatten(value))
            if list_like(address):
                address = (AddressCell(addr) for addr in flatten(address))
            else:
                address = flatten(AddressRange(address).resolve_range)
            address = tuple(address)
            assert len(address) == len(value)
            for addr, val in zip(address, value):
                self.set_value(addr, val)
            return

        elif address not in self.cell_map:
            address = AddressRange.create(address).address
            assert address in self.cell_map, (
                f'Address "{address}" not found in the cell map. Evaluate the '
                'address, or an address that references it, to place it in the cell map.')

        if set_as_range and list_like(value) and not (
                value and list_like(value[0])):
            value = (value, )

        cell_or_range = self.cell_map[address]

        if cell_or_range.value != value:  # pragma: no branch
            # need to be able to 'set' an empty cell, set to not None
            cell_or_range.value = value

            # reset the node + its dependencies
            if not self.cycles:
                self._reset(cell_or_range)

            # set the value
            cell_or_range.value = value

        # maintain dynamic array spill invariants on user writes
        if not self.cycles:
            self._note_cell_value_set(cell_or_range, address)
            self._recalc_reset_dynamic_arrays()

    def _recalc_reset_dynamic_arrays(self):
        """Re-evaluate dynamic array anchors invalidated by a set_value"""
        for a_cell in tuple(self.cell_map.values()):
            if not isinstance(a_cell, _Cell):
                continue
            if a_cell.formula is None or not a_cell.needs_calc:
                continue
            if self._is_dynamic_formula(a_cell):
                self._evaluate(str(a_cell.address))

    def _reset(self, cell):
        if cell.needs_calc:
            return

        # a resetting dynamic array anchor releases its spilled cells
        if isinstance(cell, _Cell) and str(cell.address) in self.spill_map:
            self._clear_spill(cell)

        self.log.info(f"Resetting {cell.address}")
        cell.value = None

        if cell in self.dep_graph:
            for child_cell in self.dep_graph.successors(cell):
                if child_cell.value is not None:
                    self._reset(child_cell)

    def value_tree_str(self, address, indent=0):
        iterative_eval_tracker.inc_iteration_number()
        yield from self._value_tree_str(address)

    def _value_tree_str(self, address, indent=0):
        """Generator which returns a formatted dependency graph"""
        cell = self.cell_map[address]

        if iterative_eval_tracker.is_calced(address):
            yield f'{" " * indent}{address} <- cycle'
        else:
            iterative_eval_tracker.calced(address)
            yield f'{" " * indent}{address} = {cell.value}'
            for children in sorted(self.dep_graph.predecessors(cell),
                                   key=lambda a: a.address.address):
                yield from self._value_tree_str(
                    children.address.address, indent + 1)

    def recalculate(self):
        """Recalculate all of the known cells"""
        # release all dynamic array spill regions, they are re-established
        # when their anchors are evaluated below
        for anchor_address in tuple(self.spill_map):
            anchor_cell = self.cell_map.get(anchor_address)
            if anchor_cell is not None:
                self._clear_spill(anchor_cell)

        for cell in self.cell_map.values():
            if isinstance(cell, _CellRange) or cell.formula:
                cell.value = None

        for cell in tuple(self.cell_map.values()):
            self.evaluate(cell.address.address)

    def trim_graph(self, input_addrs, output_addrs):
        """Remove unneeded cells from the graph"""
        input_addrs = tuple(AddressRange(addr).address for addr in input_addrs)
        output_addrs = tuple(AddressRange(addr) for addr in output_addrs)

        # 1) build graph for all needed outputs
        self._gen_graph(output_addrs)

        # 2) walk the dependant tree (from the inputs) and find needed cells
        needed_cells = set()

        def walk_dependents(cell):
            """passed in a _Cell or _CellRange"""
            for child_cell in self.dep_graph.successors(cell):
                child_addr = child_cell.address.address
                if child_addr not in needed_cells:
                    needed_cells.add(child_addr)
                    walk_dependents(child_cell)

        missing_dependants = set()
        for addr in input_addrs:
            try:
                if addr in self.cell_map:
                    walk_dependents(self.cell_map[addr])
                    msg = ''
                else:
                    msg = 'warning', f'Address {addr} not found in cell_map'
            except nx.exception.NetworkXError as exc:
                if AddressRange(addr) not in output_addrs:
                    msg = 'error', f'{exc}: which usually means no outputs are dependant on it.'
                else:
                    msg = 'warning', str(exc)
            if msg:
                missing_dependants.add((addr, *msg))
        if missing_dependants:
            for addr, level, msg in missing_dependants:
                getattr(self.log, level)(f"Input address {addr}: {msg}")

        if any(m[1] != 'warning' for m in missing_dependants):
            raise ValueError('\n' + '\n'.join(
                map(str, sorted(missing_dependants, key=lambda x: x[2]))))

        # even unconnected output addresses are needed
        for addr in output_addrs:
            needed_cells.add(addr.address)

        # 3) walk the precedent tree (from the output) and trim unneeded cells
        processed_cells = set()

        def walk_precedents(cell):
            for child_address in (a.address for a in cell.needed_addresses):
                if child_address not in processed_cells:  # pragma: no branch
                    processed_cells.add(child_address)
                    child_cell = self.cell_map[child_address]
                    if child_address in needed_cells or ':' in child_address:
                        walk_precedents(child_cell)
                    else:
                        # trim this cell, now we will need only its value
                        needed_cells.add(child_address)
                        child_cell.formula = None
                        self.log.debug(f'Trimming {child_address}')

        for addr in output_addrs:
            walk_precedents(self.cell_map[addr.address])

        # 4) check for any buried (not leaf node) inputs
        for addr in input_addrs:
            cell = self.cell_map.get(addr)
            if cell and getattr(cell, 'formula', None):
                self.log.info(f"{addr} is not a leaf node")

        # 5) remove unneeded cells
        cells_to_remove = tuple(addr for addr in self.cell_map
                                if addr not in needed_cells)
        for addr in cells_to_remove:
            del self.cell_map[addr]

    def validate_serialized(self, **kwargs):
        assert self.excel, "validate_serialized() needs to be run on the compiler"
        failed = self.validate_calcs(**kwargs)
        if failed != {}:
            return failed

        self.to_file(file_types='json')
        compiled = ExcelCompiler.from_file(self.filename + '.json')
        compiled.recalculate()
        for addr in self.cell_map:
            if not self.cell_map[addr].close_enough(
                    compiled.cell_map[addr].value, tol=kwargs.get('tolerance')):
                failed[addr] = Mismatch(
                    self.cell_map[addr].value, compiled.cell_map[addr].value,
                    str(self.cell_map[addr].formula))
                print('{} mismatch  {} -> {}  {}'.format(
                    addr, self.cell_map[addr].value, compiled.cell_map[addr].value,
                    str(self.cell_map[addr].formula)))
        return failed

    def validate_calcs(self, output_addrs=None, sheet=None, verify_tree=True,
                       tolerance=None, raise_exceptions=False):
        """For each address, calc the value, and verify that it matches

        This is a debugging tool which will show which cells evaluate
        differently than they do for excel.

        :param output_addrs: The cells to evaluate from (defaults to all)
        :param sheet: The sheet to evaluate from (defaults to all)
        :param verify_tree: Follow the tree to any precedent nodes
        :return: dict of addresses with good/bad values that failed to verify
        """
        if output_addrs is None:
            to_verify = list(self.formula_cells(sheet))
            print(f'Found {len(to_verify)} formulas to evaluate')
        elif list_like(output_addrs):
            to_verify = [AddressCell(addr) for addr in flatten(output_addrs)]
        else:
            to_verify = [AddressCell(output_addrs)]

        verified = set()
        failed = {}
        if self.cycles:
            iterative_eval_tracker(**self.cycles)
        while to_verify:
            addr = to_verify.pop()
            if len(to_verify) % 100 == 0:
                print(f"{len(to_verify)} formulas left to process")
            try:
                self._gen_graph(addr)
                cell = self.cell_map[addr.address]
                if isinstance(cell, _Cell) and cell.python_code and (
                        not cell.address.is_unbounded_range):
                    original_value = cell.value
                    if original_value == str(cell.formula):
                        self.log.debug(f"No Orig data?: {addr}: {cell.value}")
                        continue

                    cell.value = None
                    self.evaluate(addr.address)

                    if not (original_value is None or
                            cell.close_enough(original_value, tol=tolerance)):
                        failed.setdefault('mismatch', {})[str(addr)] = Mismatch(
                            original_value, cell.value,
                            cell.formula.base_formula)
                        print('{} mismatch  {} -> {}  {}'.format(
                            addr, original_value, cell.value,
                            cell.formula.base_formula))

                        # do it again to allow easy break-pointing
                        cell.value = None
                        self.evaluate(cell.address.address)

                verified.add(addr)
                if verify_tree:  # pragma: no branch
                    for addr in cell.needed_addresses:
                        if addr not in verified:  # pragma: no branch
                            to_verify.append(addr)
            except Exception as exc:
                if raise_exceptions:
                    raise
                cell = self.cell_map.get(addr.address, None)
                formula = cell and cell.formula.base_formula
                exc_str = str(exc)
                exc_str_split = exc_str.split('\n')

                if 'is not implemented' in exc_str:
                    exc_str_key = exc_str.split('is not implemented')[0]
                    exc_str_key = exc_str_key.strip().rsplit(' ', 1)[1].upper()
                    not_implemented = True

                else:
                    if len(exc_str_split) == 1:
                        exc_str_key = f'{type(exc).__name__}: {exc_str}'
                    else:
                        exc_str_key = exc_str_split[-2]  # pragma: no cover
                    not_implemented = exc_str_key.startswith(
                        'NotImplementedError: ')

                if not_implemented:
                    failed.setdefault('not-implemented', {}).setdefault(
                        exc_str_key, []).append((str(addr), formula, exc_str))
                else:
                    failed.setdefault('exceptions', {}).setdefault(
                        exc_str_key, []).append((str(addr), formula, exc_str))

        return failed

    def formula_cells(self, sheet=None):
        """Iterate all cells and find cells with formulas"""
        if sheet is None:
            return list(it.chain.from_iterable(
                self.formula_cells(sheet.title)
                for sheet in self.excel.workbook))

        if sheet not in self._formula_cells_dict:
            if sheet in self.excel.workbook:
                self._formula_cells_dict[sheet] = tuple(
                    AddressCell.create(cell.coordinate, sheet)
                    for row in self.excel.workbook[sheet].iter_rows()
                    for cell in row
                    if isinstance(getattr(cell, 'value', None), str) and
                    cell.value.startswith('=')
                )
            else:
                self._formula_cells_dict[sheet] = tuple()

        return self._formula_cells_dict[sheet]

    def _make_cells(self, address):
        """Given an AddressRange or AddressCell generate compiler Cells"""

        # from here don't build cells that are already in the cell_map
        assert address.address not in self.cell_map

        def add_node_to_graph(node):
            self.dep_graph.add_node(node)
            self.dep_graph.nodes[node]['sheet'] = node.sheet
            self.dep_graph.nodes[node]['label'] = node.address.coordinate

            # stick in queue to add edges
            self.graph_todos.append(node)

        def build_cell(excel_cell):
            a_cell = self.Cell(excel_cell.address, value=excel_cell.values,
                               formula=excel_cell.formula, excel=self.excel)
            self.cell_map[str(excel_cell.address)] = a_cell
            return [a_cell]

        def build_range(excel_range):
            a_range = _CellRange(excel_range, excel=self.excel)
            self.cell_map[str(excel_range.address)] = a_range

            added = [a_range]
            if isinstance(excel_range.formula, tuple):
                for addr, value, formula in a_range.cells_to_build(excel_range):
                    if addr.address not in self.cell_map:
                        a_cell = self.Cell(addr, value, formula, self.excel)
                        self.cell_map[addr.address] = a_cell
                        added.append(a_cell)
            else:
                for addr in a_range.needed_addresses:
                    if addr.address not in self.cell_map:
                        self._make_cells(addr)
            return added

        self.log.debug(f'_make_cells: {address}')
        excel_data = self.excel.get_range(address)
        if address.is_range:
            if excel_data.address != address:
                # if the actual data returned is not the same as the address
                # given, then use a reference
                self.cell_map[str(address)] = self.Cell(
                    address, formula=REF_FORMAT.format(excel_data.address),
                    excel=self.excel)

            self.range_todos.append(str(excel_data.address))
            new_nodes = build_range(excel_data)
        else:
            new_nodes = build_cell(excel_data)

        for new_node in new_nodes:
            if isinstance(new_node, _CellRange) or new_node.formula:
                # nodes to analyze: only ranges and formulas have precedents
                add_node_to_graph(new_node)

    def _evaluate_range(self, address):
        """Evaluate a range"""
        if address in ERROR_CODES:
            return address

        cell_range = self.cell_map.get(address)
        if cell_range is None:
            # we don't save the _CellRange values in the text format files
            assert '!' in address, f"{address} missing sheetname"
            self._gen_graph(address)
            cell_range = self.cell_map[address]

        if cell_range.needs_calc:
            self.log.debug(f"Evaluating: {cell_range.address}, {cell_range.python_code}")
            if cell_range.address.is_unbounded_range:
                bounded_addr = str(self.eval(cell_range))
                bounded_addr_cell = self.cell_map.get(bounded_addr)
                if bounded_addr_cell.value is None:
                    self._evaluate_range(bounded_addr)
                data = bounded_addr_cell.value

            elif cell_range.formula is None:
                data = tuple(
                    tuple(self._evaluate(addr.address) for addr in row)
                    for row in cell_range.addresses
                )
            else:
                # CSE Array Formula
                data = self.eval(cell_range, cell_range.address)
            self.log.info(f"Range {cell_range.address} evaluated to '{data}'")

            cell_range.value = data

        return cell_range.value

    def _evaluate(self, address):
        """Evaluate a single cell"""
        if address not in self.cell_map:
            # INDIRECT() and OFFSET() can produce addresses we don't already have loaded
            self._gen_graph(address)
        cell = self.cell_map[address]

        # calculate the cell value for formulas and ranges
        if cell.needs_calc:
            if isinstance(cell, _CellRange) or cell.address.is_unbounded_range:
                self._evaluate_range(cell.address.address)

            elif cell.python_code:
                self.log.debug(f"Evaluating: {address}, {cell.python_code}")
                value = self.eval(cell)
                if is_address(value):
                    # eval produced an address (aka: a reference)
                    if value.is_range:
                        # complain as we are not going to do any spilling
                        self.log.warning(f"Cell {address} evaluated to '{value}',"
                                         f" truncating to '{value.start}'")
                        value = value.start
                    else:
                        self.log.info(f"Cell {address} evaluated to address '{value}'")

                    # fetch the value for this cell, if it exists
                    ref_addr = value.address
                    if ref_addr not in self.cell_map and getattr(self, 'excel', None):
                        # INDIRECT() can produce addresses we don't already have loaded
                        self._gen_graph(ref_addr)

                    value = self.cell_map[ref_addr].value
                else:
                    self.log.info(
                        f"Cell {cell.address} evaluated to '{value}' ({type(value).__name__})")

                # a multi cell array result spills to the neighboring cells
                cell.value = self._apply_spill(cell, value)

            elif cell.formula is None:
                # an empty cell may be the spill target of a dynamic array
                # formula anchored up and to the left of it
                self._discover_spill(cell)

        return cell.value

    # ------------------------------------------------------------------
    # Dynamic array spill support (e.g. SEQUENCE)
    #
    # A formula whose final result is a rectangular array "spills": the
    # anchor cell keeps the top left element and the other elements are
    # written to the cells immediately right of and below the anchor.
    # spill_map maps anchor address -> spilled cell addresses, while
    # spill_owners is the reverse mapping.
    # ------------------------------------------------------------------

    def _apply_spill(self, cell, value):
        """Spill a multi cell formula result, return the anchor value"""
        if list_like(value):
            if not len(value):
                self._clear_spill(cell)
                cell.spill_shape = None
                return None

            if not list_like(value[0]):
                # normalize a one dimensional array to a single row
                value = (tuple(value), )

            nrows = len(value)
            ncols = len(value[0])
            if nrows == 1 and ncols == 1:
                self._clear_spill(cell)
                cell.spill_shape = None
                return value[0][0]

            return self._spill_array(cell, value, nrows, ncols)

        # a scalar result releases any previous spill
        self._clear_spill(cell)
        cell.spill_shape = None
        return value

    def _spill_array(self, cell, data, nrows, ncols):
        anchor = cell.address
        anchor_addr = str(anchor)

        # remember the desired extent even if the spill is blocked, so a
        # later edit inside the area can re-evaluate the anchor
        cell.spill_shape = (nrows, ncols)

        end_col_idx = anchor.col_idx + ncols - 1
        end_row = anchor.row + nrows - 1
        if end_col_idx > MAX_COL or end_row > MAX_ROW:
            self.log.info(f"Spill blocked for {anchor}: runs off the sheet")
            self._clear_spill(cell)
            return SPILL_ERROR

        target_rows = AddressRange(
            (anchor.col_idx, anchor.row, end_col_idx, end_row),
            sheet=anchor.sheet).resolve_range

        # the spill area (other than the anchor) must be empty
        for target in flatten(target_rows):
            target_addr = str(target)
            if target_addr == anchor_addr:
                continue
            target_cell = self._ensure_spill_cell(target_addr)
            owner = self.spill_owners.get(target_addr)
            has_formula = target_cell.formula is not None
            has_value = target_cell.value is not None
            owned_elsewhere = owner != anchor_addr
            if has_formula or (has_value and owned_elsewhere):
                self.log.info(
                    f"Spill blocked for {anchor}: {target_addr} is occupied")
                self._clear_spill(cell)
                return SPILL_ERROR

        # release any previous spill before writing the new region
        self._clear_spill(cell)

        spilled = []
        for row_idx, row in enumerate(target_rows):
            for col_idx, target in enumerate(row):
                target_addr = str(target)
                if target_addr == anchor_addr:
                    continue
                target_cell = self._ensure_spill_cell(target_addr)
                new_value = data[row_idx][col_idx]
                if target_cell.value != new_value:
                    # formulas depending on this cell must drop stale values
                    self._invalidate_dependents(target_cell)
                target_cell.spill_owner = anchor_addr
                target_cell.value = new_value
                self.spill_owners[target_addr] = anchor_addr
                spilled.append(target_addr)

        self.spill_map[anchor_addr] = tuple(spilled)
        self.log.info(f"Cell {anchor} spilled to {len(spilled)} cells")
        return data[0][0]

    def _clear_spill(self, anchor_cell):
        """Empty all cells this anchor previously spilled to"""
        anchor_addr = str(anchor_cell.address)
        spilled = self.spill_map.pop(anchor_addr, ())
        for target_addr in spilled:
            self.spill_owners.pop(target_addr, None)
            target_cell = self.cell_map.get(target_addr)
            if target_cell is None:
                continue
            target_cell.spill_owner = None
            if target_cell.value is not None:
                # reset so formulas depending on the spilled value follow
                self._reset(target_cell)

    def _ensure_spill_cell(self, target_addr):
        if target_addr not in self.cell_map:
            self._gen_graph(target_addr)
        return self.cell_map[target_addr]

    def _invalidate_dependents(self, cell):
        """Reset formulas with cached values that depend on this cell"""
        if cell in self.dep_graph:
            for child_cell in self.dep_graph.successors(cell):
                if child_cell.value is not None:
                    self._reset(child_cell)

    def _discover_spill(self, cell):
        """Find a dynamic array anchor which spills onto this empty cell"""
        target_addr = str(cell.address)
        if target_addr in self.spill_owners:
            return

        for anchor_addr in self._spill_candidates(cell.address):
            if anchor_addr in self._spill_resolving:
                continue
            self._spill_resolving.add(anchor_addr)
            try:
                self._evaluate(anchor_addr)
            finally:
                self._spill_resolving.discard(anchor_addr)
            if target_addr in self.spill_owners:
                return

    def _spill_candidates(self, target):
        """Dynamic array formula cells above-left of `target`"""
        candidates = {}

        def add(a_cell):
            if a_cell is None or a_cell.formula is None:
                return
            address = a_cell.address
            if address.sheet != target.sheet:
                return
            if address.row > target.row or address.col_idx > target.col_idx:
                return
            anchor_addr = str(address)
            if anchor_addr == str(target):
                return
            if self._is_dynamic_formula(a_cell):
                candidates[anchor_addr] = (
                    target.row - address.row,
                    target.col_idx - address.col_idx,
                )

        for a_cell in self.cell_map.values():
            if isinstance(a_cell, _Cell):
                add(a_cell)

        # dynamic formula cells may not have been built into the cell map
        # yet, look directly in the underlying workbook when available
        excel = getattr(self, 'excel', None)
        workbook = getattr(excel, 'workbook', None)
        if workbook is not None and target.sheet in workbook.sheetnames:
            for anchor_addr in self._workbook_dynamic_anchors(target.sheet):
                formula_cell = AddressRange(anchor_addr).start
                if formula_cell.row > target.row:
                    continue
                if formula_cell.col_idx > target.col_idx:
                    continue
                if anchor_addr not in self.cell_map:
                    self._gen_graph(anchor_addr)
                add(self.cell_map.get(anchor_addr))

        # nearest anchor first
        return tuple(sorted(candidates, key=lambda a: candidates[a]))

    def _workbook_dynamic_anchors(self, sheet):
        """Addresses of dynamic array formulas in a workbook sheet

        Workbook formulas never change over the lifetime of a compiler,
        so the scan is cached.  The sparse openpyxl ``_cells`` mapping is
        used when available: merely accessing a far away cell inflates
        ``max_row`` and makes a plain ``iter_rows`` scan touch millions
        of empty rows.
        """
        cached = self._dynamic_anchor_cache.get(sheet)
        if cached is not None:
            return cached

        worksheet = self.excel.workbook[sheet]
        sparse_cells = getattr(worksheet, '_cells', None)
        if sparse_cells is not None:
            cells = sparse_cells.values()
        else:
            cells = (opxl_cell for row in worksheet.iter_rows()
                     for opxl_cell in row)

        anchors = tuple(
            str(AddressCell.create(opxl_cell.coordinate, sheet=sheet))
            for opxl_cell in cells
            if isinstance(opxl_cell.value, str) and
            opxl_cell.value.startswith('=') and
            'sequence(' in opxl_cell.value.lower()
        )
        self._dynamic_anchor_cache[sheet] = anchors
        return anchors

    @staticmethod
    def _is_dynamic_formula(cell):
        formula = cell.formula
        if formula is None:
            return False
        source = formula.base_formula or formula.python_code or ''
        return 'sequence(' in source.lower()

    def _note_cell_value_set(self, cell, address):
        """Maintain spill invariants after a direct set_value()"""
        if not isinstance(cell, _Cell):
            return

        address = str(cell.address)

        if address in self.spill_map or cell.spill_shape is not None:
            # writing the anchor itself releases the spill, the anchor
            # just holds the written value
            self._clear_spill(cell)
            cell.spill_shape = None
            return

        # writing into a currently spilled cell turns it into user data,
        # i.e. an obstacle for the anchor
        owner = self.spill_owners.pop(address, None)
        if owner is not None:
            cell.spill_owner = None
            remaining = tuple(
                a for a in self.spill_map.get(owner, ()) if a != address)
            if remaining:
                self.spill_map[owner] = remaining
            else:
                self.spill_map.pop(owner, None)

        # re-evaluate anchors whose (possibly blocked) spill covers here,
        # this both creates the #SPILL! and restores it when cleared
        for anchor_addr in self._anchors_covering(cell.address):
            anchor_cell = self.cell_map.get(anchor_addr)
            if anchor_cell is None:
                continue
            if not anchor_cell.needs_calc:
                self._reset(anchor_cell)
            if anchor_cell.needs_calc:
                self._evaluate(anchor_addr)

    def _anchors_covering(self, target):
        anchors = []
        for anchor_addr, anchor_cell in self.cell_map.items():
            if not isinstance(anchor_cell, _Cell):
                continue
            shape = anchor_cell.spill_shape
            if not shape:
                continue
            address = anchor_cell.address
            if address.sheet != target.sheet:
                continue
            nrows, ncols = shape
            within_rows = address.row <= target.row <= address.row + nrows - 1
            within_cols = (
                address.col_idx <= target.col_idx <= address.col_idx + ncols - 1)
            if within_rows and within_cols:
                anchors.append(anchor_addr)
        return tuple(anchors)

    def _restore_spills(self, saved):
        """Restore spill ownership after loading a serialized compiler"""
        # evaluation during graph construction may already have spilled,
        # reset that bookkeeping before recreating it from saved data
        for spilled in self.spill_map.values():
            for target_addr in spilled:
                target_cell = self.cell_map.get(target_addr)
                if target_cell is not None:
                    target_cell.spill_owner = None
        self.spill_map = {}
        self.spill_owners = {}

        if not saved:
            return

        for anchor_addr, spilled in saved.items():
            anchor_cell = self.cell_map.get(anchor_addr)
            if anchor_cell is None or anchor_cell.formula is None:
                continue

            owners = []
            max_row_off = max_col_off = 0
            anchor_row = anchor_cell.address.row
            anchor_col = anchor_cell.address.col_idx
            for target_addr in spilled:
                target_cell = self.cell_map.get(target_addr)
                if target_cell is None:
                    continue
                target_cell.spill_owner = anchor_addr
                self.spill_owners[target_addr] = anchor_addr
                owners.append(target_addr)
                max_row_off = max(
                    max_row_off, target_cell.address.row - anchor_row)
                max_col_off = max(
                    max_col_off, target_cell.address.col_idx - anchor_col)

            if owners:
                self.spill_map[anchor_addr] = tuple(owners)
                anchor_cell.spill_shape = (max_row_off + 1, max_col_off + 1)

    def _evaluate_non_iterative(self, address):
        """ evaluate a cell or cells in the spreadsheet

        :param address: str, AddressRange, AddressCell or a tuple or list
            or iterable of these three
        :return: evaluated value/values
        """
        if str(address) not in self.cell_map:
            if list_like(address):
                if not isinstance(address, (tuple, list)):
                    address = tuple(address)

                # process a tuple or list of addresses
                return type(address)(
                    self._evaluate_non_iterative(c) for c in address)

            address = AddressRange.create(address)

            # get the sheet if not specified
            if not address.has_sheet:
                address = AddressRange(
                    address, sheet=self.excel.get_active_sheet_name())

            if address.address not in self.cell_map:
                self._gen_graph(address)

        result = self._evaluate(str(address))
        if isinstance(result, tuple):
            # trim excess dimensions
            if len(result[0]) == 1:
                result = tuple(row[0] for row in result)
            if len(result) == 1:
                result = result[0]
        return result

    def _evaluate_iterative(self, address, iterations=None, tolerance=None):
        """ evaluate a cell or cells in a spreadsheet with cycles

        reference: https://support.microsoft.com/en-us/office/
                    8540bd0f-6e97-4483-bcf7-1b49cd50d123

        :param address: str, AddressRange, AddressCell or a tuple or list
            or iterable of these three
        :param iterations: maximum number of iterations to compute. If not
            specified use the value from the workbook.
        :param tolerance: maximum change, if any calculated value changes by
            more than this, another iteration will be performed. If not
            specified use the value from the workbook.
        :return: evaluated value/values
        """

        iterations = iterations or self.cycles['iterations'] or 10000
        tolerance = tolerance or self.cycles['tolerance'] or 0.01

        progress_tracker = iterative_eval_tracker(iterations, tolerance)
        while True:
            progress_tracker.inc_iteration_number()
            results = self._evaluate_non_iterative(address)
            if progress_tracker.done:
                return results

    def _gen_graph(self, seed, recursed=False):
        """Given a starting point (e.g., A6, or A3:B7) on a particular sheet,
        generate a Spreadsheet instance that captures the logic and control
        flow of the equations.
        """
        if not is_address(seed):
            if isinstance(seed, str):
                seed = AddressRange(seed)
            elif isinstance(seed, collections.abc.Iterable):
                for s in seed:
                    self._gen_graph(s, recursed=True)
                self._process_gen_graph()
                return
            else:
                raise ValueError(f'Unknown seed: {seed}')

        # get/set the current sheet
        if not seed.has_sheet:
            seed = AddressRange(seed, sheet=self.excel.get_active_sheet_name())

        if '[' in seed.sheet:
            raise NotImplementedError('Linked SheetNames')

        if seed.address in self.cell_map:
            # already did this cell/range
            return

        # process the seed
        self._make_cells(seed)

        if not recursed:
            # if not entered to process one cell / cellrange process other work
            self._process_gen_graph()

    def _process_gen_graph(self):

        while self.graph_todos:
            # connect the dependant cells in the graph
            dependant = self.graph_todos.pop()

            self.log.debug(f"Handling {dependant.address}")

            for precedent_address in dependant.needed_addresses:
                if precedent_address.address not in self.cell_map:
                    self._gen_graph(precedent_address, recursed=True)

                self.dep_graph.add_edge(
                    self.cell_map[precedent_address.address], dependant)

        # calc the values for ranges
        try:
            for range_todo in reversed(self.range_todos):
                self._evaluate_range(range_todo)
        finally:
            self.range_todos = []

        self.log.info(
            f"Graph construction done, {len(self.dep_graph.nodes())} nodes, "
            f"{len(self.dep_graph.edges())} edges, "
            f"{len(self.cell_map)} self.cell_map entries"
        )

    def eval_conditional_formats(self, address):
        """Evaluate the conditional format (formulas) for a cell or cells

        returns the conditional format id which is the key for the dict:
          ExcelCompiler.conditional_formats

        NOTE: conditional_formats are not saved in the persistent formats.
              If needed they can be hand serialized into "extra_data"

        :param address: str, AddressRange, AddressCell or a tuple or list
            or iterable of these three
        :return: evaluated objects ids
        """
        if list_like(address):
            if not isinstance(address, (tuple, list)):
                address = tuple(address)

            # process a tuple or list of addresses
            return type(address)(
                self.eval_conditional_formats(c) for c in address)

        address = AddressRange.create(address)

        # get the sheet if not specified
        if not address.has_sheet:
            address = AddressRange(
                address, sheet=self.excel.get_active_sheet_name())

        if address.is_range:
            return tuple(tuple(self.eval_conditional_formats(addr)
                               for addr in row) for row in address.rows)

        cf_addr = str(address).replace('!', '.cf!')

        if cf_addr not in self.cell_map:
            phony_cell = _Cell(address)
            formats = self.excel.conditional_format(address)
            format_strs = []
            for f in formats:
                excel_formula = ExcelFormula(f.formula, cell=phony_cell)
                python_code = excel_formula.python_code
                format_strs.append(
                    f'({python_code}, {f.dxf_id}, {int(bool(f.stop_if_true))})')
                self.conditional_formats[f.dxf_id] = f.dxf

            python_code = f"=conditional_format_ids({', '.join(format_strs)})"
            a_cell = _Cell(address, formula=python_code)
            self.cell_map[cf_addr] = a_cell
            self._gen_graph(a_cell.formula.needed_addresses)

        return self.eval(self.cell_map[cf_addr])


class _CellBase:

    value = None

    # dynamic array bookkeeping:
    #  spill_owner: anchor address spilling into this cell, if any
    #  spill_shape: (rows, cols) of the last array produced by an anchor
    spill_owner = None
    spill_shape = None

    def __init__(self, address=None, formula='', excel=None):
        formula_is_python_code = excel is None or isinstance(
            excel, _CompiledImporter)
        self.formula = formula and ExcelFormula(
            formula, cell=self,
            formula_is_python_code=formula_is_python_code) or None

        if isinstance(excel, _CompiledImporter):
            excel = None
        self.excel = excel
        self.address = AddressRange(address)

    @property
    def sheet(self):
        return self.address.sheet

    @property
    def python_code(self):
        return self.formula and self.formula.python_code

    @property
    def needs_calc(self):
        return self.value is None

    def close_enough(self, value, rel=0.00001, tol=None):
        if isinstance(self.value, Number) and isinstance(value, Number):
            if tol is not None:
                return abs(value - self.value) < (1 + rel) * tol
            elif value and self.value:
                return math.isclose(self.value, value, rel_tol=rel)
            else:
                return math.isclose(self.value, value, abs_tol=1e-8)
        else:
            return self.value == value


class _CellRange(_CellBase):
    # TODO: only supports rectangular ranges

    def __init__(self, data, excel=None):
        formula = None
        if data.formula and isinstance(data.formula, str):
            formula = data.formula
            if formula.startswith('={') and formula[-1] == '}':
                formula = '=' + formula[2:-1]

        super().__init__(address=data.address, formula=formula, excel=excel)
        if not self.address.sheet:
            raise ValueError(f"Must pass in a sheet: {self.address}")

        self.addresses = data.address.resolve_range
        self.size = data.address.size
        self.value = None

    def __getstate__(self):
        state = dict(self.__dict__)
        state['excel'] = None
        return state

    def __repr__(self):
        return str(self.address)

    __str__ = __repr__

    def __iter__(self):
        return flatten(self.addresses)

    @property
    def serialize(self):
        # Ranges with formulas need to be serialized
        return bool(self.formula)

    @property
    def needed_addresses(self):
        return self.formula and self.formula.needed_addresses or iter(self)

    def cells_to_build(self, data):
        assert isinstance(data.formula, tuple)
        return zip(  # pragma: no branch
            self,  # address
            flatten(v for row in data.values for v in row),  # value
            flatten(f for row in data.formula for f in row)  # formula
        )


class _Cell(_CellBase):
    ctr = 0
    serialize = True

    @classmethod
    def next_id(cls):
        cls.ctr += 1
        return cls.ctr

    def __init__(self, address, value=None, formula='', excel=None):
        super().__init__(address=address, formula=formula, excel=excel)

        self.value = value

        # every cell has a unique id
        self.id = _Cell.next_id()

    def __getstate__(self):
        state = dict(self.__dict__)
        state['excel'] = None
        return state

    def __repr__(self):
        return f"{self.address} -> {self.formula or self.value}"

    __str__ = __repr__

    @property
    def needed_addresses(self):
        return self.formula and self.formula.needed_addresses or ()


class _CycleCell(_Cell):
    """Cell which participates in a iterative calculation

    For non iterative (non-cyclic) excel sheets we use reset() (set value
    to None), then calc anything that is None.  But for iterative (cyclic)
    excel sheets the inputs to a cell could potentially change anytime, so
    we need to calc everything all the time.

    While it would be possible to break things done further to cyclic and
    non-cyclic sections of the graph, and then only iterate on the cyclic
    sections, implemented here is a simpler algorithm:

    1. Start at the top of eval tree (cell to evaluate)
    2. Mark the cell in question as being a work in progress (WIP)
    3. Eval (ie: calc the lambda for) the cell.  The will cause other
       cells to be evaluated
    4. If the value of a  cell that is WIP is needed, then we have a loop.
       Use the previous value, and do not descend any farther on the tree.
    5. After evaluating a cell, check if the value changed by more that
       the allowed tolerance, if so note the cell as needing more evals
    """

    def __init__(self, *args, **kwargs):
        self._value = None
        self._prev_value = None
        self.wip = False
        super().__init__(*args, **kwargs)

    @property
    def value(self):
        if self.wip:
            return self._prev_value
        else:
            return self._value

    @value.setter
    def value(self, a_value):
        iterative_eval_tracker.calced(self)
        self.wip = False
        self._value = a_value
        if not self.close_enough(
                self._prev_value, tol=iterative_eval_tracker.tolerance):
            iterative_eval_tracker.wip(self)

    def start_calcs(self):
        self.wip = True
        self._prev_value = self._value

    @property
    def needs_calc(self):
        return not self.wip and not iterative_eval_tracker.is_calced(self)


class _CompiledImporter:
    """Emulate the excel_wrapper for serialized files"""
    def __init__(self, filename, file_data):
        # take the workbook filename from the deserialized data if available
        # otherwise the passed in filename
        self.filename = file_data.get(
            'filename', filename.rsplit('.', maxsplit=1)[0])
        self.cell_map = file_data['cell_map']
        self.compiler = None

    def get_range(self, address):
        cell = self._get_cell(address)

        if not address.is_range:
            return cell

        elif address.is_unbounded_range:
            # this is a unbounded range to range mapping, disassemble
            formula = cell.formula
            assert formula.startswith(REF_START)
            assert formula.endswith(REF_END)
            ref_addr = formula[len(REF_START):-len(REF_END)]
            return self.get_range(AddressRange(ref_addr))

        elif cell.formula:
            return cell

        else:
            # need to map col or row ranges to a specific range
            addresses = address.resolve_range

            cells = [[self._get_cell(addr) for addr in row]
                     for row in addresses]
            values = [[c.values for c in row] for row in cells]

            return ExcelOpxWrapper.RangeData(address, None, values)

    def _get_cell(self, address):
        cell_value = self.cell_map.get(str(address))

        if cell_value is None:
            return ExcelOpxWrapper.RangeData(address, '', None)

        elif isinstance(cell_value, str) and cell_value.startswith('='):
            return ExcelOpxWrapper.RangeData(address, cell_value, None)

        else:
            return ExcelOpxWrapper.RangeData(address, '', cell_value)

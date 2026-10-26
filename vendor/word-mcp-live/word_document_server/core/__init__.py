"""Backward-compatible lazy exports."""
from importlib import import_module
_EXPORTS = {'ensure_heading_style': ('word_document_server.core.styles', 'ensure_heading_style'), 'ensure_table_style': ('word_document_server.core.styles', 'ensure_table_style'), 'create_style': ('word_document_server.core.styles', 'create_style'), 'add_protection_info': ('word_document_server.core.protection', 'add_protection_info'), 'verify_document_protection': ('word_document_server.core.protection', 'verify_document_protection'), 'is_section_editable': ('word_document_server.core.protection', 'is_section_editable'), 'create_signature_info': ('word_document_server.core.protection', 'create_signature_info'), 'verify_signature': ('word_document_server.core.protection', 'verify_signature'), 'add_footnote': ('word_document_server.core.footnotes', 'add_footnote'), 'add_endnote': ('word_document_server.core.footnotes', 'add_endnote'), 'convert_footnotes_to_endnotes': ('word_document_server.core.footnotes', 'convert_footnotes_to_endnotes'), 'find_footnote_references': ('word_document_server.core.footnotes', 'find_footnote_references'), 'get_format_symbols': ('word_document_server.core.footnotes', 'get_format_symbols'), 'customize_footnote_formatting': ('word_document_server.core.footnotes', 'customize_footnote_formatting'), 'set_cell_border': ('word_document_server.core.tables', 'set_cell_border'), 'apply_table_style': ('word_document_server.core.tables', 'apply_table_style'), 'copy_table': ('word_document_server.core.tables', 'copy_table')}
def __getattr__(name):
    if name not in _EXPORTS: raise AttributeError(name)
    module, member = _EXPORTS[name]
    value = getattr(import_module(module), member)
    globals()[name] = value
    return value

"""Semantic enumerations used throughout the public API.

These enums describe Access concepts in Python terms. They are deliberately *not* the raw COM constants
(``acTextBox = 109`` etc.); the COM adapters translate between the two, so no magic numbers ever appear in
user code. Every enum is a :class:`~enum.StrEnum`, which keeps specs readable when serialized to JSON/YAML.
"""

from __future__ import annotations

from enum import StrEnum

__all__ = [
    "ControlKind",
    "DataType",
    "DialogPolicy",
    "Engine",
    "FormView",
    "JoinType",
    "LayoutKind",
    "MacroSecurity",
    "ModuleKind",
    "NumberSize",
    "ObjectKind",
    "PropertyType",
    "QueryKind",
    "RowSourceType",
    "ScrollBars",
    "Section",
    "Transport",
]


class Engine(StrEnum):
    """Which automation engine a session may use (see ``AccessDatabase.open(engine=...)``)."""

    AUTO = "auto"
    """In-process DAO when available, otherwise Access; upgrades to Access for design features."""
    DAO = "dao"
    """In-process DAO only. Never starts ``MSACCESS.EXE``; design features raise ``CapabilityError``."""
    ACCESS = "access"
    """Microsoft Access automation (works with any Python/Office bitness combination)."""


class Transport(StrEnum):
    """How a session currently reaches the database (reported by ``AccessDatabase.transport``)."""

    DAO_INPROC = "dao-inproc"
    """DAO loaded into the Python process."""
    ACCESS_HOSTED = "access-hosted"
    """DAO hosted by a hidden Access instance; the database is not opened in the Access UI."""
    ACCESS_DESIGN = "access-design"
    """The database is Access's current database; forms, reports and modules are available."""
    MEMORY = "memory"
    """In-memory fake used for tests and dry runs."""


class DataType(StrEnum):
    """Access field data types, named after the Access table designer."""

    TEXT = "text"
    """Short Text (up to 255 characters)."""
    LONG_TEXT = "long_text"
    """Long Text (formerly Memo)."""
    NUMBER = "number"
    """Number; the storage size is given by :class:`NumberSize`."""
    DECIMAL = "decimal"
    """Number with Field Size = Decimal (fixed precision and scale)."""
    CURRENCY = "currency"
    AUTONUMBER = "autonumber"
    DATE_TIME = "date_time"
    YES_NO = "yes_no"
    HYPERLINK = "hyperlink"
    OLE_OBJECT = "ole_object"
    ATTACHMENT = "attachment"
    CALCULATED = "calculated"
    LARGE_NUMBER = "large_number"
    DATE_TIME_EXTENDED = "date_time_extended"
    BINARY = "binary"
    """Fixed/variable binary (not creatable from the Access UI)."""
    MULTI_VALUE = "multi_value"
    """A multi-valued (complex) lookup field."""
    UNKNOWN = "unknown"
    """A type PyAccessKit does not model yet; the raw DAO type is preserved on introspection."""


class NumberSize(StrEnum):
    """The *Field Size* of a Number field, exactly as offered by the Access table designer.

    Note that Access's ``INTEGER`` is a **16-bit** integer; the default, ``LONG_INTEGER``, is 32-bit.
    """

    BYTE = "byte"
    INTEGER = "integer"
    LONG_INTEGER = "long_integer"
    SINGLE = "single"
    DOUBLE = "double"
    REPLICATION_ID = "replication_id"


class JoinType(StrEnum):
    """Default join type of a relationship (as shown in the Relationships window)."""

    INNER = "inner"
    LEFT = "left"
    """Include all records from the primary table."""
    RIGHT = "right"
    """Include all records from the foreign (related) table."""


class QueryKind(StrEnum):
    """Kind of a saved query (derived from DAO ``QueryDef.Type``)."""

    SELECT = "select"
    CROSSTAB = "crosstab"
    DELETE = "delete"
    UPDATE = "update"
    APPEND = "append"
    MAKE_TABLE = "make_table"
    DDL = "ddl"
    PASS_THROUGH = "pass_through"
    UNION = "union"
    PASS_THROUGH_BULK = "pass_through_bulk"
    COMPOUND = "compound"
    PROCEDURE = "procedure"
    ACTION = "action"
    UNKNOWN = "unknown"


class ObjectKind(StrEnum):
    """Kinds of objects stored in an Access database."""

    TABLE = "table"
    QUERY = "query"
    FORM = "form"
    REPORT = "report"
    MACRO = "macro"
    MODULE = "module"
    # Sub-objects used in error messages.
    FIELD = "field"
    INDEX = "index"
    RELATIONSHIP = "relationship"
    PROPERTY = "property"
    CONTROL = "control"
    DATABASE = "database"


class ModuleKind(StrEnum):
    """VBA module kinds."""

    STANDARD = "standard"
    CLASS = "class"


class FormView(StrEnum):
    """A form's *Default View*."""

    SINGLE = "single"
    CONTINUOUS = "continuous"
    DATASHEET = "datasheet"
    SPLIT = "split"


class ScrollBars(StrEnum):
    """A form's *Scroll Bars* setting."""

    NEITHER = "neither"
    HORIZONTAL = "horizontal"
    VERTICAL = "vertical"
    BOTH = "both"


class Section(StrEnum):
    """Form sections a control can be placed in."""

    DETAIL = "detail"
    HEADER = "header"
    FOOTER = "footer"


class LayoutKind(StrEnum):
    """How a form's controls are arranged when they have no explicit position."""

    AUTO = "auto"
    """``STACKED`` for single/split forms, ``TABULAR`` for continuous and datasheet forms."""
    STACKED = "stacked"
    """One control per row with its label on the left (like Access's stacked layout)."""
    TABULAR = "tabular"
    """Labels in the form header and one row of controls in the detail section."""
    NONE = "none"
    """No automatic layout: every control must be positioned with ``at=``."""


class RowSourceType(StrEnum):
    """Where a combo box gets its rows from."""

    TABLE_QUERY = "table_query"
    """A table name, query name or SQL statement."""
    VALUE_LIST = "value_list"
    """A semicolon-separated list of values."""


class ControlKind(StrEnum):
    """Kinds of form controls (for building and introspection)."""

    LABEL = "label"
    TEXTBOX = "textbox"
    CHECKBOX = "checkbox"
    COMBOBOX = "combobox"
    LISTBOX = "listbox"
    BUTTON = "button"
    OPTION_GROUP = "option_group"
    OPTION_BUTTON = "option_button"
    TOGGLE_BUTTON = "toggle_button"
    SUBFORM = "subform"
    IMAGE = "image"
    LINE = "line"
    RECTANGLE = "rectangle"
    TAB_CONTROL = "tab_control"
    PAGE = "page"
    ATTACHMENT = "attachment"
    OTHER = "other"


class MacroSecurity(StrEnum):
    """What Access may run when PyAccessKit opens a database as the current database."""

    DISABLE = "disable"
    """Disable all macros and VBA (``msoAutomationSecurityForceDisable``). Default; deterministic."""
    USE_UI = "ui"
    """Follow the user's Trust Center settings (``msoAutomationSecurityByUI``)."""
    ENABLE = "enable"
    """Enable all macros and VBA, including AutoExec (``msoAutomationSecurityLow``)."""


class DialogPolicy(StrEnum):
    """What to do when Access shows a modal dialog during an automated call."""

    FAIL = "fail"
    """Dismiss the dialog and raise ``AccessDialogError`` (default)."""
    WARN = "warn"
    """Dismiss the dialog and emit an ``AccessDialogWarning``."""
    OFF = "off"
    """Do not watch for dialogs (a blocking dialog can then hang until ``call_timeout``)."""


class PropertyType(StrEnum):
    """Data types for user-defined DAO properties (``CreateProperty``)."""

    BOOLEAN = "boolean"
    BYTE = "byte"
    INTEGER = "integer"
    LONG = "long"
    CURRENCY = "currency"
    SINGLE = "single"
    DOUBLE = "double"
    DATE_TIME = "date_time"
    TEXT = "text"
    MEMO = "memo"

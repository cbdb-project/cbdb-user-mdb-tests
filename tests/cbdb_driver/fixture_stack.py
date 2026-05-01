"""
fixture_stack — declarative scenario steps for CBDB form testing.

Each Step encapsulates one mutation or assertion. A Stack runs them in
order against a FormDriver. Stacks compose: a higher-level test can
re-use a "preamble" stack and append its own steps.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterable

from .form_driver import FormDriver


@dataclass
class Step:
    """Base. Subclasses override ``apply``. ``label`` shows in failure
    traces."""
    label: str = ""

    def apply(self, drv: FormDriver) -> None:  # pragma: no cover
        raise NotImplementedError


# ---------- mutators ----------

@dataclass
class OpenForm(Step):
    form: str = ""
    hidden: bool = True
    def __post_init__(self):
        self.label = self.label or f"OpenForm({self.form})"
    def apply(self, drv):
        drv.open_form(self.form, hidden=self.hidden)


@dataclass
class CloseForm(Step):
    form: str = ""
    def __post_init__(self):
        self.label = self.label or f"CloseForm({self.form})"
    def apply(self, drv):
        drv.close_form(self.form)


@dataclass
class SetControl(Step):
    form: str = ""
    ctl: str = ""
    value: Any = None
    def __post_init__(self):
        self.label = self.label or f"SetControl({self.form}.{self.ctl}={self.value!r})"
    def apply(self, drv):
        drv.set_control(self.form, self.ctl, self.value)


@dataclass
class GetControl(Step):
    form: str = ""
    ctl: str = ""
    def __post_init__(self):
        self.label = self.label or f"GetControl({self.form}.{self.ctl})"
    def apply(self, drv):
        return drv.get_control(self.form, self.ctl)


@dataclass
class SetGlobal(Step):
    var: str = ""
    value: Any = None
    def __post_init__(self):
        self.label = self.label or f"SetGlobal({self.var}={self.value!r})"
    def apply(self, drv):
        drv.set_global(self.var, self.value)


@dataclass
class ClearTable(Step):
    table: str = ""
    def __post_init__(self):
        self.label = self.label or f"ClearTable({self.table})"
    def apply(self, drv):
        drv.app.exec_sql(f"DELETE FROM [{self.table}]")


@dataclass
class SetPickerCodes(Step):
    """Replace contents of a picker scratch table with the given ID list."""
    table: str = ""
    ids: Iterable[int] = ()
    column: str = "c_entry_code"
    def __post_init__(self):
        self.label = self.label or f"SetPickerCodes({self.table}={list(self.ids)!r})"
    def apply(self, drv):
        drv.app.exec_sql(f"DELETE FROM [{self.table}]")
        for i in self.ids:
            drv.app.exec_sql(
                f"INSERT INTO [{self.table}] ([{self.column}]) VALUES ({int(i)})"
            )


@dataclass
class SetPickerAddrs(Step):
    """Convenience for ZZ_SCRATCH_ADDR (column = c_addr_id)."""
    addr_ids: Iterable[int] = ()
    table: str = "ZZ_SCRATCH_ADDR"
    def __post_init__(self):
        self.label = self.label or f"SetPickerAddrs({list(self.addr_ids)!r})"
    def apply(self, drv):
        drv.app.exec_sql(f"DELETE FROM [{self.table}]")
        for a in self.addr_ids:
            drv.app.exec_sql(
                f"INSERT INTO [{self.table}] (c_addr_id) VALUES ({int(a)})"
            )


@dataclass
class InvokeEvent(Step):
    form: str = ""
    event: str = ""
    def __post_init__(self):
        self.label = self.label or f"InvokeEvent({self.form}.{self.event})"
    def apply(self, drv):
        drv.invoke_event(self.form, self.event)


@dataclass
class InvokeEventUnchecked(Step):
    form: str = ""
    event: str = ""
    def __post_init__(self):
        self.label = self.label or f"InvokeEventUnchecked({self.form}.{self.event})"
    def apply(self, drv):
        drv.invoke_event_unchecked(self.form, self.event)


# ---------- assertions ----------

@dataclass
class AssertEnabled(Step):
    form: str = ""
    ctl: str = ""
    expected: bool = True
    def __post_init__(self):
        self.label = self.label or f"AssertEnabled({self.form}.{self.ctl}={self.expected})"
    def apply(self, drv):
        actual = bool(drv.get_control_property(self.form, self.ctl, "Enabled"))
        if actual != self.expected:
            raise AssertionError(
                f"{self.form}.{self.ctl}.Enabled = {actual}, expected {self.expected}"
            )


@dataclass
class AssertRowCount(Step):
    table: str = ""
    expected: int | None = None
    min_rows: int | None = None
    max_rows: int | None = None
    where: str | None = None
    def __post_init__(self):
        self.label = self.label or f"AssertRowCount({self.table}, expected={self.expected}, min={self.min_rows}, max={self.max_rows})"
    def apply(self, drv):
        n = drv.app.row_count(self.table, self.where)
        if self.expected is not None and n != self.expected:
            raise AssertionError(f"{self.table} count = {n}, expected {self.expected}")
        if self.min_rows is not None and n < self.min_rows:
            raise AssertionError(f"{self.table} count = {n}, expected >= {self.min_rows}")
        if self.max_rows is not None and n > self.max_rows:
            raise AssertionError(f"{self.table} count = {n}, expected <= {self.max_rows}")


@dataclass
class AssertNoErrors(Step):
    """Explicitly assert ZZ_TEST_ERRORS is empty (mainly defensive — every
    InvokeEvent already does this implicitly)."""
    label: str = "AssertNoErrors"
    def apply(self, drv):
        n = drv.app.row_count("ZZ_TEST_ERRORS")
        if n > 0:
            errs = drv.app.fetch_all(
                "SELECT form_name, event_name, err_desc FROM ZZ_TEST_ERRORS ORDER BY id"
            )
            lines = "\n".join(f"  [{r.form_name}.{r.event_name}] {r.err_desc}" for r in errs)
            raise AssertionError(f"unexpected VBA errors logged:\n{lines}")


# ---------- runner ----------

class Stack:
    def __init__(self, steps: list[Step]):
        self.steps = list(steps)

    def __add__(self, other) -> "Stack":
        if isinstance(other, Stack):
            return Stack(self.steps + other.steps)
        if isinstance(other, list):
            return Stack(self.steps + other)
        if isinstance(other, Step):
            return Stack(self.steps + [other])
        return NotImplemented

    def run(self, drv: FormDriver) -> None:
        for i, s in enumerate(self.steps):
            try:
                s.apply(drv)
            except AssertionError as e:
                raise AssertionError(
                    f"step #{i+1} ({s.label}) failed:\n  {e}"
                ) from e
            except Exception as e:
                raise AssertionError(
                    f"step #{i+1} ({s.label}) raised {type(e).__name__}: {e}"
                ) from e

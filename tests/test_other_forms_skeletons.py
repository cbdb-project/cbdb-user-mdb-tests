"""
Skeleton tests for the other 9 LookAt forms.

Each form needs:
  1. A replay module under tests/cbdb_replay/  (use TEMPLATE_lookat.py
     as a starting point — port the CmdQuery_Click logic from
     analysis/dump/vba/Form_LookAt<Formname>.vb)
  2. A test file like test_lookatentry.py with parameterised fixtures
     and assert_matches_golden checks.

Until those are written, the tests below act as a smoke test: they
verify the result table the form populates exists and is empty
(its expected starting state in a fresh copy of the .mdb).
"""
from __future__ import annotations

import pytest


# (form_name, result_table)
LOOKAT_FORMS = [
    ("LookAtAssociations",     "ZZ_SCRATCH_ASSOC"),
    ("LookAtAssociationPairs", "ZZ_SCRATCH_PAIR_PEOPLE"),
    ("LookAtKinship",          "ZZ_SCRATCH_KIN"),
    ("LookAtNetworks",         "ZZ_SOCIAL_NETWORK"),
    ("LookAtOffice",           "ZZ_SCRATCH_OFFICE"),
    ("LookAtPlace",            "ZZ_SCRATCH_PLACE_PEOPLE"),
    ("LookAtStatus",           "ZZ_SCRATCH_STATUS"),
    ("LookAtTexts",            "ZZ_SCRATCH_BIOG_TEXT_DATA"),
    ("LookAtGroupData",        "ZZ_SCRATCH_PEOPLE"),
]


@pytest.mark.parametrize("form_name,result_table", LOOKAT_FORMS,
                         ids=[f for f, _ in LOOKAT_FORMS])
def test_result_table_exists(ro_conn, form_name, result_table):
    """The form's result/scratch table is present in the schema."""
    cur = ro_conn.cursor()
    cur.execute(f"SELECT TOP 1 * FROM [{result_table}]")
    cur.fetchall()  # OK if zero rows
    cur.close()


@pytest.mark.skip(reason="TODO: implement replay modules for these forms")
@pytest.mark.parametrize("form_name,result_table", LOOKAT_FORMS,
                         ids=[f for f, _ in LOOKAT_FORMS])
def test_replay_implemented(form_name, result_table):
    """Placeholder.  Remove the @skip and implement once the
    corresponding replay module exists in cbdb_replay/."""
    pytest.fail(
        f"replay for {form_name} not implemented; see "
        "tests/cbdb_replay/TEMPLATE_lookat.py for the recipe"
    )

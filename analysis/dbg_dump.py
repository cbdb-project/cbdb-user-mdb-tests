"""Dump ZZ_TEST_DEBUG from a working-copy mdb."""
import sys, win32com.client as wc
mdb = sys.argv[1] if len(sys.argv) > 1 else r"C:\Users\how612\Documents\GitHub\cbdb-user-mdb-tests\analysis\_bug_behavior_test_copy.mdb"
app = wc.DispatchEx("Access.Application")
app.Visible = False
app.OpenCurrentDatabase(mdb)
db = app.CurrentDb()
rs = db.OpenRecordset("SELECT id, msg FROM ZZ_TEST_DEBUG ORDER BY id ASC")
i = 0
while not rs.EOF and i < 60:
    print(f"{rs.Fields(0).Value}: {rs.Fields(1).Value}")
    rs.MoveNext(); i += 1
app.CloseCurrentDatabase(); app.Quit()

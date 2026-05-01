"""Try Dispatch + DispatchEx and see which works."""
import gc, time, win32com.client, subprocess, sys
print(f"python {sys.version}")
print(f"win32com {win32com.__file__}")
sys.stdout.flush()

subprocess.run(["taskkill","/F","/IM","MSACCESS.EXE"],
               capture_output=True, check=False)
time.sleep(1)

print("Trying Dispatch...", flush=True)
t0 = time.time()
try:
    app = win32com.client.Dispatch("Access.Application")
    print(f"  Dispatch OK in {time.time()-t0:.2f}s", flush=True)
    try: app.Quit()
    except Exception: pass
    del app
    gc.collect()
except Exception as e:
    print(f"  Dispatch FAILED in {time.time()-t0:.2f}s: {e}", flush=True)

subprocess.run(["taskkill","/F","/IM","MSACCESS.EXE"],
               capture_output=True, check=False)
time.sleep(1)

print("Trying DispatchEx...", flush=True)
t0 = time.time()
try:
    app = win32com.client.DispatchEx("Access.Application")
    print(f"  DispatchEx OK in {time.time()-t0:.2f}s", flush=True)
    try: app.Quit()
    except Exception: pass
    del app
    gc.collect()
except Exception as e:
    print(f"  DispatchEx FAILED in {time.time()-t0:.2f}s: {e}", flush=True)

print("done", flush=True)

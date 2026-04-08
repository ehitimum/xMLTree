import subprocess
import time
import signal
import sys

print("Starting xMLTree.py in subprocess...")
proc = subprocess.Popen([sys.executable, "src/xMLTree.py"], 
                        stdout=subprocess.PIPE, 
                        stderr=subprocess.PIPE,
                        text=True)

# Wait a few seconds for app to start
time.sleep(3)

# Check if process is still running
if proc.poll() is None:
    print("Application started successfully (process still running)")
    # Send SIGTERM to terminate
    proc.terminate()
    try:
        stdout, stderr = proc.communicate(timeout=2)
        print("Stdout:", stdout[:500])
        print("Stderr:", stderr[:500])
    except subprocess.TimeoutExpired:
        proc.kill()
        stdout, stderr = proc.communicate()
        print("Process killed after timeout")
else:
    # Process exited early
    stdout, stderr = proc.communicate()
    print("Process exited with code:", proc.returncode)
    print("Stdout:", stdout)
    print("Stderr:", stderr)
    sys.exit(1)

print("Test completed successfully")

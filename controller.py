import subprocess
import time

# Run CSV merging
subprocess.run(["python", "merge.py"])

# Wait a bit for file to save
time.sleep(2)

# Run macro script
subprocess.run(["python", "run_macro.py"])
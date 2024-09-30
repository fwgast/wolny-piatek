import sys, glob, os
power_automate_path = os.path.join(os.environ['ProgramFiles'], 'WindowsApps', 'Microsoft.PowerAutomateDesktop_*', 'PAD.Console.Host.exe')
EXE_PATH = glob.glob(power_automate_path)
sys.exit(EXE_PATH)
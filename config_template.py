# -*- coding: utf-8 -*-
from os import path, environ
from glob import glob
SUBJECT = 'subject'
SENDER = 'sender@mail.com'
power_automate_path = path.join(environ['ProgramFiles'], 'WindowsApps', 'Microsoft.PowerAutomateDesktop_*', 'PAD.Console.Host.exe')
EXE_PATH = glob(power_automate_path)
PA_LINK = """ zbedny link bo to funkcja premium :)) ale zostawiam to moze kiedys to bedzie za free """
BODY_SENSITIVE_DATA = """ BODY_SENSITIVE_DATA """
FOOTER = """
FOOTER
"""
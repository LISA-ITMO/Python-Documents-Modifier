import os
import sys

for x in os.walk('~/src'):
  sys.path.insert(0, x[0])
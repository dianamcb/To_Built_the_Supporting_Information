# Note: Execute this script only once (unless it prints any error), otherwise the files will be
#   re-renamed.

import os, re

# List all files in the path of the scrpt.
fullPaths = []
for root, dirs, files in os.walk("./"):
    for file in files: fullPaths.append(os.path.join(root, file))

# Rename files according to their level of theory.
# Path format: .../SPE-[Level of theory]-[A, P, r or T]...Pathway...
# A, P, r and T come from Anionic, Protonic, reactants and TMS, respectively.
for i in range(len(fullPaths)):
    renamedFullPath = re.sub('(.+)(SPE-.+)(-[APrT].+Pathway.+)(\..+)', r'\1\2\3-\2\4', fullPaths[i])
    # print(renamedFullPath)
    os.rename(fullPaths[i], renamedFullPath)


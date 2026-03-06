import pandas as pd
import re
import subprocess
import time
import warnings
import os
import glob

warnings.simplefilter("ignore")

start_time=time.time()

print("Traitement des exports NCE ...")

scripts= ['reparation.py','logbrut.py','parc.py','ExpGPON.py','ExpXDSL.py','ExpSP.py','espacedisque.py','main.py']

#for script in scripts:

 
# #print(f"Execution du script de traitement {script}...")

subprocess.run(['python',scripts[0]])

subprocess.run(['python',scripts[1]])

subprocess.run(['python',scripts[2]])

subprocess.run(['python',scripts[3]])

subprocess.run(['python',scripts[4]])

subprocess.run(['python',scripts[5]])

subprocess.run(['python',scripts[6]])

subprocess.run(['python',scripts[7]])


print(f"Fin de l'execution du controle NCE")

for _ in range(1000000):
    pass
end_time=time.time()

elapsed_time=end_time-start_time

min = elapsed_time // 60
sec= elapsed_time % 60


print(f"Temps d'execution : {int(min)} minutes et {sec:.2f} secondes")
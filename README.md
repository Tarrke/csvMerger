# csvMerger

un script qui extrait des cellules d'un csv pour en faire une ligne dans un excel résultat, pour tous les csv d'un répertoire.

## Utilisation

~~~bash
$ ./csvMerger.py -h
usage: csvMerger.py [-h] [-d DIR] [-v] [-f CONFIG] output

A simple python script that fetch data into csvs and write an excel output
file

positional arguments:
  output                output file name

optional arguments:
  -h, --help            show this help message and exit
  -d DIR, --dir DIR     directory where to fetch data
  -v, --verbose         increase output verbosity
  -f CONFIG, --config CONFIG
                        configuration file for the parser
~~~

## Todo

* ~~ gestion des inputs de la ligne de commande ~~
* utilisation d'un fichier de configuration
* listing des fichiers csv d'un répertoire
* extraction des cellules demandées
* applicaiton d'un algorithme sur les cellules demandées pour transformation
* écriture d'un fichier de sortie  excel

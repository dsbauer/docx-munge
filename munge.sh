#!/bin/bash
set -e

#Ensure starting state:
#1) Make a folder/directory named NAME for each NAME.docx file

#2) Move NAME.docx within directory NAME (e.g. IVC3.docx will be in directory IVC3/)

#3) Collect all linked documents cited by NAME.docx in the NAME directory

#4) Each linked document should be named NAME-N-whatever, for some integer N
#    (e.g. IVC3-11-fuzzy_kittens.pdf)
#  The citation NAME-N should be the name of the link in NAME.docx referencing
#  the original remote document (http://somewhere/fuzzyKittens.pdf)

#5) cd to the directory containing all document folders (IVC1/, IVC2/, etc)
#  The command `ls` should then list 'IVC1 IVC2 IVC3'...

#6) Set the location where the new links will start looking for documents.
#  Run the command `export SDCCD_PATH=/Some/Local/Directory`
#  Then the links in NAME.docx will point to files in /Some/Local/Directory/NAME.
#  That location may be different from the current directory, but it will eventually
#  need to have the same document subdirectories (e.g. IVC1, IVC2...)
#  This command only needs to be run once and will persist in that command terminal
#  as long as the terminal window remains open.

#  As often as needed...
#7) Run this script with the command `./munge.sh NAME` where NAME is IVC3, IVC4, etc

# in commands below, $1 holds NAME (e.g. IVC3)
mkdir $1/_unzipped
cd $1/_unzipped
unzip ../$1.docx #unpack the docx components
node ../../munge.js $1 #transform the hyperlinks
#rm word/_rels/.document.xml.rels #uncomment to remove unused backup file
zip -r ../munged-$1.docx * #rebundle the components into new docx
cd ..
#rm -r _unzipped
echo "Created file $1/munged-$1.docx"

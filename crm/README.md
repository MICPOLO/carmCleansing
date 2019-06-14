To run this java program, you must pass the file path as program argument

The program will do the following tasks
1.Remove tabular content in word documents
2.Recursively extract the OLE file 
3.Extract specific tabular content
4.Label the title and sub-title

The configuration of the above tasks is under the src/main/resources/config.prperties file

config.prperties file

the ole file to be extracted
whiteList 

the embedded image to be kept
EmbedFormatException 

the keyword of tabular content 
rowException = Purpose

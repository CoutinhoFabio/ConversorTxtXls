This project is intended to create excel xlsx file using data from a text file with tab-separated values and vice versa.
To convert from text file to xlsx the content of file should follow this structure:
- each line in text file is equivalent to each row of xlsx sheet
- each tab char will put the next data in the next col of xlsx sheet
Example converting text file to xlsx:

[text file content]
Hello#9World

[calling converter in cmd prompt]
java -jar MavenXLS.java TXT2XLS HelloWorld.txt HelloWorld.xlsx
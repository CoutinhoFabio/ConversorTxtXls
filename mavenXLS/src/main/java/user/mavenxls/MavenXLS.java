/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Project/Maven2/JavaApp/src/main/java/${packagePath}/${mainClassName}.java to edit this template
 */

package galli.mavenxls;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import java.io.FileWriter;
import java.util.Iterator;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import org.apache.commons.io.IOUtils;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFCreationHelper;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
/**
 *
 * @author Usuario
 */
public class MavenXLS {
    private static final boolean TESTE = false;
    
    
    public static boolean isFilenameValid(String file) 
    {
        File f = new File(file);
        try 
        {
            f.getCanonicalPath();
            return true;
        }
        catch (IOException e) 
        {
            return false;
        }
    }
   
    public static boolean isNumeric(String str) 
    {
        return str.matches("-?\\d+(\\.\\d+)?");  //match a number with optional '-' and decimal.
    }
    
    public static void main(String[] args) {
        boolean txt2xls = false;
        int lineNumber;
        String txtFile = "";
        String xlsFile = "";
        
        //Lemos todos os args passados através de um foreach
        for (String arg : args) 
        {
        
            //captura o argumento
            arg = arg.toUpperCase().trim();
        
            if (arg.equals("TXT2XLS"))
            {
                txt2xls= true;   
            }
            else 
            if (arg.equals("XLS2TXT"))
            {
                txt2xls= false;   
            }
            else
            if(arg.contains(".XLS"))
            {
                xlsFile= arg; 
            }
            else 
            if(arg.contains(".TXT"))
            {
                txtFile= arg;  
            }
        }
    
        
        if(TESTE)
        {
            txt2xls= true;
            txtFile= "E:\\projetos_java\\mavenXLS\\target\\Test1.txt";
            xlsFile= "test1.xls";
        }
        
        if(txtFile.isEmpty())
        {
            System.out.print("error: text file name");
            return;
            
        }
        else
        if(xlsFile.isEmpty())
        {
            System.out.print("error: xls file name");
            return;
        }
        
        //agora que pegou os argumentos, trabalha com eles
    
    
        if(txt2xls == true)
        {
            //operação de escrita (txt --> excel)
        
         
            //read file into stream, try-with-resources
            try
            {
                int rowCount = 0;
                
                File f = new File(txtFile);
                if(!f.exists() || f.isDirectory()) 
                { 
                    // do something
                    System.out.print("error: unable to read text file");
                    return;
                }
            
                //XSSFWorkbook workbook = new XSSFWorkbook();
                //XSSFSheet sheet = workbook.createSheet("Datalogger");
                HSSFWorkbook workbook = new HSSFWorkbook();
                HSSFSheet sheet = workbook.createSheet("Datalogger");
                
                Scanner scanner= new Scanner(f);
                scanner.useDelimiter("\r\n");
                lineNumber= 1;
                while(scanner.hasNext())
                {
              
                    //pega o conteúdo de uma linha do arquivo de texto
                    String input = scanner.next();
                
                    //separa palavra por palavra da linha, separados por "tab"
                    String[] words= input.split("\t");
                
                    //cria uma linha no arquivo excel
                    Row row = sheet.createRow(rowCount++);
                    int columnCount = 0;
                    for( int i = 0; i <= words.length - 1; i++)
                    {
                
                        Cell cell = row.createCell(columnCount++);
                    
                        //pega um item da linha do arquivo texto
                        String item = words[i];
                    
                        //verifica o comando da linha do arquivo texto
                        if(isFilenameValid(item) && item.contains(".bmp"))
                        {
                            InputStream inputStream = new FileInputStream(item);
                            
                            //Get the contents of an InputStream as a byte[].
                            byte[] bytes = IOUtils.toByteArray(inputStream);
                            
                            //Adds a picture to the workbook
                            
                            int pictureIdx = workbook.addPicture(bytes, HSSFWorkbook.PICTURE_TYPE_PNG);
                            
                            //close the input stream
                            inputStream.close();
                            
                            //Returns an object that handles instantiating concrete classes
                            HSSFCreationHelper helper = workbook.getCreationHelper();
                            
                            //Creates the top-level drawing patriarch.
                            HSSFPatriarch drawing = sheet.createDrawingPatriarch();
 
                            //Create an anchor that is attached to the worksheet
                            HSSFClientAnchor anchor = helper.createClientAnchor();

                            //create an anchor with upper left cell _and_ bottom right cell
                            anchor.setRow1(rowCount-1);
                            anchor.setCol1(columnCount-1);
   
                            anchor.setRow2(rowCount);
                            anchor.setCol2(columnCount);

                            //Creates a picture
                            HSSFPicture pict = drawing.createPicture(anchor, pictureIdx);

                            //Reset the image to the original size
                            //pict.resize(); //don't do that. Let the anchor resize the image!

                            //Create the Cell B3
                            cell = sheet.createRow(rowCount).createCell(columnCount);

                            //set width to n character widths = count characters * 256
                            int widthUnits = 20*256;
                            sheet.setColumnWidth(1, widthUnits);

                            //set height to n points in twips = n * 20
                            short heightUnits = 60*20;
                            cell.getRow().setHeight(heightUnits);
                        }
                        else
                        if(isNumeric(item))
                        {
                            //adiciona coluna do arquivo excel
                            double val= Double.parseDouble(item);
                            cell.setCellValue(val);
                        }
                        else
                        {
                            //adiciona coluna do arquivo excel
                            cell.setCellValue(item);
                        }
                    }
                
                    lineNumber++; //contador de linhas do arquivo texto
                }
            
                FileOutputStream outputStream = new FileOutputStream(xlsFile);
                workbook.write(outputStream);
            }
            catch(Exception e)
            {
                System.out.print("error: converting text file");
                return;
            }
        
            //try (FileOutputStream outputStream = new FileOutputStream(xlsFile)) 
            //{
            //    workbook.write(outputStream);
           // }
            //catch(Exception e)
            //{
            //    System.out.print("error: unable to save xls file");
            //    return;
            //}
        }
        else
        {
            //operação de leitura (excel --> txt)
            
            try
            {
                //OPCPackage file = OPCPackage.open(new FileInputStream(xlsFile));
                HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(new File(xlsFile)));
                //FileInputStream file = new FileInputStream(new File(xlsFile));
 
                //Create Workbook instance holding reference to .xlsx file
                //XSSFWorkbook workbook = new XSSFWorkbook(file);
 
                //Get first/desired sheet from the workbook
                HSSFSheet sheet = workbook.getSheetAt(0);
 
                FileWriter destFile = new FileWriter(txtFile);
                
                //Iterate through each rows one by one
                Iterator<Row> rowIterator = sheet.iterator();
                while (rowIterator.hasNext()) 
                {
                    Row row = rowIterator.next();
                    //For each row, iterate through all the columns
                    Iterator<Cell> cellIterator = row.cellIterator();
                 
                    while (cellIterator.hasNext()) 
                    {
                        Cell cell = cellIterator.next();
                        
                        //Check the cell type and format accordingly
                        switch (cell.getCellType()) 
                        {
                            case 0: //CELL_TYPE_NUMERIC:
                                destFile.write(cell.getNumericCellValue() + "\t");
                            break;
                            case 1: //CELL_TYPE_STRING:
                                destFile.write(cell.getStringCellValue() + "\t");
                            break;
                        }
                    }
                    destFile.write("\r\n");
                }
                //file.close();
                destFile.close();
            } 
            catch (Exception e) 
            {
                System.out.print(e.getMessage());
                //System.out.print("error: unable to convert xls to text file");
                return;
            }
        }

        System.out.print("0");
    }
}

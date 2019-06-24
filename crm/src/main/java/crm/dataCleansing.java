package crm;
import java.awt.Color;
import java.io.File;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;


import org.apache.tika.cli.TikaCLI;
import org.apache.tika.exception.TikaException;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.Parser;
import org.apache.tika.parser.txt.TXTParser;
import org.apache.tika.sax.BodyContentHandler;
import org.xml.sax.SAXException;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Font;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.OoxmlSaveOptions;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import com.aspose.words.SaveFormat;
import com.aspose.words.Shape;
import com.google.common.io.Files;

import org.apache.tika.parser.AutoDetectParser;
import org.apache.commons.io.comparator.LastModifiedFileComparator;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.util.Properties;

public class dataCleansing {
	
	static ArrayList<String> whiteList;
	static ArrayList<String> EmbedFormatException;
	static ArrayList<String> rowException;
	static Boolean keepTable;
	
	
	public static void main(String[] args) throws Exception {
		
		Properties p =new Properties();
		InputStream is = new FileInputStream(new File("src/main/resources/config.properties"));
		p.load(is);
		//whitelist = files format to be kept
		whiteList = new ArrayList<String>(Arrays.asList(p.getProperty("whiteList").toString().split(",")));
		//EmbedFormatException = OLE object to be kept
		EmbedFormatException = new ArrayList<String>(Arrays.asList(p.getProperty("EmbedFormatException").toString().split(",")));
		//rowException = Extract text from the table
		rowException = new ArrayList<String>(Arrays.asList(p.getProperty("rowException").toString().split(",")));
		//keepTable true = convert table to text with tag, false = remove table
		keepTable = Boolean.parseBoolean(p.getProperty("keepTable").toString());
		is.close();
		
		String sourcePath = args[0];
		File sourceRTF = new File(sourcePath);		  
		
		// parentFolder/sample.rtf
		String fileNameExtension = sourceRTF.getName();																		//sample.rtf 
		String fileEx = sourceRTF.getName().substring(sourceRTF.getName().lastIndexOf('.')+1,sourceRTF.getName().length());	//rtf
		String filePath = sourceRTF.getParentFile().getPath();																//parentFolder
		String fileName = fileNameExtension.substring(0,fileNameExtension.lastIndexOf('.'));								//sample
		String fileLoc = filePath+'/'+fileName;																				//parentFolder/sample
		
		File dir = new File(fileLoc);
		dir.mkdir();
		
		File attDir = new File(fileLoc+'/'+"Attachment");
		attDir.mkdir();
		
		File initFile = new File(fileLoc+'/'+"Attachment"+'/'+fileNameExtension);
		Files.copy(sourceRTF,initFile);
		
		initRecursiveExtract(initFile);
		
		
		
		if(fileEx.equals("docx")){
			//sample.docx
			//sample_cleansed_output.docx
			String orignalFileName =fileLoc+'/'+fileNameExtension; 
			String cleansedFileName =fileLoc+'/'+fileName+"_cleansed_output"+".docx";
			
			File originalSourceRTF = new File(orignalFileName);
			Files.copy(sourceRTF,originalSourceRTF );
			
			File newSourceRTF = new File(cleansedFileName);
			Files.copy(sourceRTF,newSourceRTF );
						
			mainHandler(newSourceRTF.getPath());
			
		}
		
		if(fileEx.equals("doc")||fileEx.equals("rtf")){
			//sample.rtf
			//sample.docx
			//sample_cleansed_output.docx
			String originalFileName=fileLoc+'/'+fileNameExtension; 
			String cleansedFileName=fileLoc+'/'+fileName+"_cleansed_output"+".docx";
			String docxFileName   =fileLoc+'/'+fileName+".docx";
			
			File originalSourceRTF = new File(originalFileName);
			Files.copy(sourceRTF,originalSourceRTF );
			
			OoxmlSaveOptions saveOptions= new OoxmlSaveOptions();
			saveOptions.setSaveFormat(SaveFormat.DOCX);
			
			File docxSourceFile = new File(docxFileName);
			Files.copy(sourceRTF,docxSourceFile );
			Document doc = new Document(docxFileName);		
			doc.save(docxFileName, saveOptions);
			
			
			Document doc2 = new Document(sourceRTF.getPath());		
			doc2.save(cleansedFileName, saveOptions);
			File newChildFile = new File(cleansedFileName);
			
			mainHandler(newChildFile.getPath());
			
		}
		System.out.println("Finish"); 
	}
	
	public static void mainHandler(String fileName) throws Exception{

		extractTrName(fileName);
		if(rowException!=null){
		insertTableRowToText(rowException,fileName);
		}
		mainEntityLabel(fileName);
		removeNoiseEmbedFile(fileName);
		if(keepTable==true){
			convertTableToText(fileName);	
		}else{
			docxDelTable(fileName);
		}

		
		
	}
	
	public static void extractTrName(String filePath) throws Exception{
		XWPFDocument XWPFdoc = new XWPFDocument(new FileInputStream(filePath));
		String trName = "";
		boolean isNext = false;
		List <XWPFTable> table = XWPFdoc.getTables();
		for (XWPFTable xwpfTable : table) {
			List<XWPFTableRow> row = xwpfTable.getRows();
			for (XWPFTableRow xwpfTableRow : row) {
				List<XWPFTableCell> cell = xwpfTableRow.getTableCells();
				if(isNext == true){
					break;
				}
				for (XWPFTableCell xwpfTableCell : cell) {
					if(xwpfTableCell.getText().contains("TR NAME") )
					{
						isNext = true; 
						continue;
					}
					if(isNext == true && !xwpfTableCell.getText().trim().equals(":") &&!xwpfTableCell.getText().trim().equals("")){
						trName = xwpfTableCell.getText();
						trName.replace(":","");
						break;
					}
				}
				
			}
			
		}
		XWPFdoc.close();
		if(isNext==true){
			String outputTrName ="<trName>";
			outputTrName += trName;
			outputTrName += "<trName>";
			Document doc = new Document(filePath);		
			DocumentBuilder builder = new DocumentBuilder(doc);
			Font font = builder.getFont();
			font.setSize(10);
			font.setBold(false);
			font.setName("Calibri");
			font.setColor(Color.black);
			builder.write(outputTrName);
			doc.save(filePath);
		
		}
		System.out.println("trName Extraction Completed");
		
	}
	
	public static void convertTableToText(String filePath) throws FileNotFoundException, IOException{
		String tableContent ="";
		XWPFDocument XWPFdoc = new XWPFDocument(new FileInputStream(filePath));
		List <XWPFTable> table = XWPFdoc.getTables();
		for(XWPFTable allTable:table){
			tableContent ="";
			XWPFTable xwpfTable =allTable;
			List<XWPFTableRow> row = xwpfTable.getRows();
			int lastRow = row.size();
			int rowCounter = 0;
			
			for (XWPFTableRow xwpfTableRow : row) {
				String rowText ="";
				rowCounter++;
				List<XWPFTableCell> cell = xwpfTableRow.getTableCells();
				tableContent +="*";
				for (XWPFTableCell xwpfTableCell : cell) {
					if((xwpfTableCell.getText().trim().length()>0))
					{
						tableContent+="|";
						tableContent+=xwpfTableCell.getText();
					}
				}
				if(rowCounter!=row.size())
				tableContent+="\r\n";
			}
			tableContent+="|";
			
			XWPFTableRow finalRow = xwpfTable.insertNewTableRow(0);
			while(xwpfTable.getNumberOfRows()>1){
				xwpfTable.removeRow(1);	
				
			}
			XWPFTableRow firstRow = xwpfTable.getRow(0);
			firstRow.addNewTableCell();
			XWPFTableCell firstCol = firstRow.getCell(0);
			firstCol.setText(tableContent);

			xwpfTable.setWidth("100%");
		}
		
		FileOutputStream out = new FileOutputStream( new File( filePath ) );
        XWPFdoc.write( out );
        out.close();
        
		System.out.println("Table to text completed");
		
	}
	
	public static void mainEntityLabel(String fileName) throws Exception {
		Document doc = new Document(fileName);
		String pattern = "(^\\d\\.)";
		Pattern regexPattern = Pattern.compile("^\\d\\.");
		int counter = 0;
		String Title = "";
		
		
		NodeCollection paragraphs = doc.getChildNodes(NodeType.PARAGRAPH,true);
		for(Paragraph paragraph :(Iterable<Paragraph>)paragraphs){
			if(paragraph.getAncestor(NodeType.TABLE)==null)
			{
				
			String runText ="";
			
			NodeCollection runs = paragraph.getRuns();
			int boldLast = 0;
			int boldCurrent = 0;
			
			boolean isList =false;
			
			for(Run run :(Iterable<Run>)runs){
				if(run.getFont().getBold()){
					boldLast++;
				}
				Matcher matcher = regexPattern.matcher(run.getText());
				if(run.getParentParagraph().isListItem()||matcher.find()){
					isList=true;
				}
			}
			if(boldLast>0){
			System.out.println("Last"+boldLast);
			for(Run run :(Iterable<Run>)runs){	 
				if(run.getFont().getBold()){
				boldCurrent++;
				System.out.println("Current"+boldCurrent);	
				//Title
				if(isList){
					if(boldCurrent==boldLast){
						counter++;
						runText+= run.getText();
						if(runText.contains("  ")){
							String[] splitedText = runText.split("  ",2);
							Matcher matcher = regexPattern.matcher(splitedText[0]);
							if(matcher.find()){
								splitedText[0] = matcher.replaceAll("");
							}else{
								run.getParentParagraph().getListFormat().removeNumbers();
							}
							Title = splitedText[0];
							run.setText("<header>"+counter+". "+splitedText[0]+"<header>"+" "+splitedText[1]);
						}
						else{
						runText = runText.trim().replaceAll("(\\s)+", " ");
						Matcher matcher = regexPattern.matcher(runText);
						if(matcher.find()){
							runText = matcher.replaceAll("");
						}else{
							run.getParentParagraph().getListFormat().removeNumbers();
						}						
						Title = runText;
						run.setText("<header>"+counter+". "+runText+"<header>");
						}
					}else{
						runText+= run.getText();	
						run.setText("");
					}
					
				}else if(counter>0){
					//Subtitle
					if(boldCurrent==boldLast && titleChecker(run)){
						runText+= run.getText();
						runText = runText.trim().replaceAll("(\\s)+", " ");
						run.setText("<header>"+counter+". "+Title+" -> "+runText+"<header>");
					}else if(titleChecker(run)){
						runText+= run.getText();		
						run.setText("");
					}					
				}
				
			}
		}
		}
				
			
			}
		}
		
		doc.save(fileName);
		
	}
	

	public static boolean titleChecker(Run run){
		boolean isTitle = true;
		if(run.getFont().getColor().equals(Color.BLUE)){
			isTitle = false;
			System.out.println("Not black - blue");
		}
		
		if(run.getFont().getColor().equals(Color.RED)){
			isTitle = false;
			System.out.println("Not black -red");
		}
		
		if(!(run.getText().trim().length()>0)){
			isTitle = false;
			System.out.println("All space");
		}	
		if(run.getFont().getItalic()){
			isTitle = false;
			System.out.println("Italic");
		}
		if(run.getFont().getName().toString().equals("Wingdings")){
			isTitle = false;
			System.out.println("Symbol");
		}
		
		return isTitle;
	}
	
	public static void removeNoiseEmbedFile(String filename) throws Exception{
		
		Document doc = new Document(filename);
		NodeCollection shapes =doc.getChildNodes(NodeType.SHAPE, true);
		for(Shape shape:(Iterable<Shape>) shapes){
			if (shape.getOleFormat() != null){
				if(!EmbedFormatException.contains(shape.getOleFormat().getProgId())){
				doc.getChildNodes(NodeType.SHAPE, true).remove(shape);
				}
			}else{
				doc.getChildNodes(NodeType.SHAPE, true).remove(shape);
			}
			
		}
		
		doc.save(filename);
		System.out.println("Noise embedded file is removed");
		
	}
	
	
	public static String getExtension(String filename) {
        if (filename == null) {
            return null;
        }
        int extensionPos = filename.lastIndexOf('.');
        
        return filename.substring(extensionPos+1,filename.length());
        
    }
	
	public static File newFile(File f, String newName,String extension) {
	
		  return new File(f.getParent() + "/" + newName + "." + extension);
		}
	
	public static File changeExtension(File f, String newExtension) {
		  int i = f.getName().lastIndexOf('.');
		  String name = f.getName().substring(0,i);
		  return new File(f.getParent() + "/" + name + newExtension);
		}
	
	//docx start
		//remove table
		private static void docxDelTable(String filename){
			File docxFile = new File(filename);	
			try {
	            FileInputStream in = new FileInputStream( new File( filename ) );
	            XWPFDocument document = new XWPFDocument( in );

	            showTablesInfo( document );

	            // Deleting the first table of the document
	            deleteTable( document, 0 );

	            showTablesInfo( document );
	            
	            saveDoc( document, docxFile.getPath() );

	        } catch ( FileNotFoundException e ) {
	            System.out.println( "File " + filename + " not found." );
	        } catch ( IOException e ) {
	            System.out.println( "IOException while processing file " + filename + ":\n" + e.getMessage() );
	        }
		}
		
	    private static void showTablesInfo( XWPFDocument document ) {
	        List<XWPFTable> tables = document.getTables();
	        System.out.println( "\n document has " + tables.size() + " table(s)." );

	        for ( XWPFTable table : tables ) {
	            System.out.println( "table with position #" + document.getPosOfTable( table ) + " has "
	                    + table.getRows().size() + " rows" );
	        }
	    }

	    private static void deleteTable( XWPFDocument document, int tableIndex ) {
	        try {
	        	while(true){
	            int bodyElement = getBodyElementOfTable( document, tableIndex );
	            System.out.println( "deleting table with bodyElement #" + bodyElement );
	            document.removeBodyElement( bodyElement );
	        	}
	        } catch ( Exception e ) {
	            System.out.println( "There is no table #" + tableIndex + " in the document." );
	        }
	    }

	    private static int getBodyElementOfTable( XWPFDocument document, int tableNumberInDocument ) {
	        List<XWPFTable> tables = document.getTables();
	        XWPFTable theTable = tables.get( tableNumberInDocument );

	        return document.getPosOfTable( theTable );
	    }

	    private static void saveDoc( XWPFDocument document, String filename ) {
	        try {
	            FileOutputStream out = new FileOutputStream( new File( filename ) );
	            document.write( out );
	            out.close();
	        } catch ( FileNotFoundException e ) {
	            System.out.println( e.getMessage() );
	        } catch ( IOException e ) {
	            System.out.println( "IOException while saving to " + filename + ":\n" + e.getMessage() );
	        }
	    }
	    
	    //docx end
	    
	    //Extract exception row
	    public static void insertTableRowToText(ArrayList<String> rowException,String filePath) throws Exception {
			
	    	ArrayList<String> rowEx = rowException;
			String tableTarget ="";
			XWPFDocument XWPFdoc = new XWPFDocument(new FileInputStream(filePath));
			List <String> rowArray = new ArrayList<String>();
			List <XWPFTable> table = XWPFdoc.getTables();
			for (XWPFTable xwpfTable : table) {
				List<XWPFTableRow> row = xwpfTable.getRows();
				for (XWPFTableRow xwpfTableRow : row) {
					String rowText ="";
					List<XWPFTableCell> cell = xwpfTableRow.getTableCells();
					for (XWPFTableCell xwpfTableCell : cell) {
						if(xwpfTableCell!=null)
						{
							rowText+= xwpfTableCell.getText()+" ";
						}
					}
					rowArray.add(rowText);
				}
			}
			XWPFdoc.close();
			
			for(String test:rowArray){
				for(String Ex: rowEx)
					if(test.contains(Ex)){
						tableTarget+="<"+Ex+">";
						tableTarget+= test;
						tableTarget+="<"+Ex+">";
						tableTarget+="\n";
						break;
					}
			}
			
	    	if(!tableTarget.equals("")){
				Document doc = new Document(filePath);		
				DocumentBuilder builder = new DocumentBuilder(doc);
				Font font = builder.getFont();
				font.setSize(10);
				font.setBold(false);
				font.setName("Calibri");
				font.setColor(Color.black);
				builder.write(tableTarget);
				doc.save(filePath);
			
				
			
				System.out.println("Exception row is extracted");
	
				System.out.println("["+tableTarget+"]");
	    	}
		}
	

		public static void initRecursiveExtract(File initFile) throws Exception{
			ArrayList<String> vbaList = new ArrayList<String>(Arrays.asList("doc","docx","rtf"));
			ArrayList<String> imageList = new ArrayList<String>(Arrays.asList("emf","wmf","jpg","jpeg","png"));
			String filePath = initFile.getPath().substring(0,initFile.getPath().lastIndexOf("\\"));
			String fileLoc = filePath;
			String input = initFile.getPath();
			//create an empty file to store modified file and embedded files
			//fileName as the folder name
			File dir = new File(fileLoc);	
			dir.mkdir();
			
			//extract all embedded doc in initFile
			try{
		        String[] arguments = new String[]{"-z", "--extract-dir="+fileLoc, input};
		        System.out.println("Using TIKA CLI to dedect embedded Files. Target Directory: "+ fileLoc);
		        TikaCLI.main(arguments);
		        
		    }
		    catch(Exception e){
		    	System.out.println("Exception in saveEmbedds, during search in File: " + input + "\r\nDetails: " + e);
		    }
			
			initFile.delete();
			//get the imageName
			ArrayList<String> imageName = new ArrayList<String>(); 
			File [] directoryListing = dir.listFiles();
			Arrays.sort(directoryListing, LastModifiedFileComparator.LASTMODIFIED_COMPARATOR);	
			if(directoryListing!=null){
				for(File child :directoryListing){
					//get doc to rtf
					String extension = child.getName().substring(child.getName().lastIndexOf('.')+1,child.getName().length());
					//before delete useless file,
					//extract the name and rename file
					if(extension.equals("emf")){
						Parser parser = new AutoDetectParser();
						BodyContentHandler handler = new BodyContentHandler();
						Metadata metadata = new Metadata();   //empty metadata object 
						FileInputStream inputstream = new FileInputStream(child);
						ParseContext context = new ParseContext();
						parser.parse(inputstream, handler, metadata, context);

						// now this metadata object contains the extracted metadata of the given file.
						String childHandler = handler.toString();
						String childName = childHandler.substring(0,childHandler.lastIndexOf('.'));
						imageName.add(childName.trim().replaceAll("\\s","_"));
						inputstream.close();
					}							
				}
			}
			//delete image
			if(directoryListing!=null){
				for(File child :directoryListing){
					String extension = child.getName().substring(child.getName().lastIndexOf('.')+1,child.getName().length());
					if(imageList.contains(extension)){
						File file = new File(fileLoc +'/'+ child.getName());
						boolean isDelete = file.delete();
						System.out.println(fileLoc +'/'+ child.getName()+" is deleted from directory");
						System.out.println(isDelete);
					}
				}
			}
			//rename file
			//the file order of second layer
			directoryListing = dir.listFiles();
			Arrays.sort(directoryListing, LastModifiedFileComparator.LASTMODIFIED_COMPARATOR);		
			if(directoryListing!=null){
				int childIndex=0;
				for(File child :directoryListing){
					String extension = child.getName().substring(child.getName().lastIndexOf('.')+1,child.getName().length());
					if(!imageList.contains(extension)){
					System.out.println("File index: "+childIndex);
					File newChild = newFile(child,imageName.get(childIndex),extension);		
					System.out.println(child.getName()+" is renamed to "+ newChild.getName());
					boolean isSuccess = child.renameTo(newChild);
					System.out.println(isSuccess);
					childIndex++;
					}
				}
			}
			//delete non-whiteList doc
			directoryListing = dir.listFiles();
			if(directoryListing!=null){
				for(File child :directoryListing){
					String extension = child.getName().substring(child.getName().lastIndexOf('.')+1,child.getName().length());
					if(!whiteList.contains(extension)){
						File file = new File(fileLoc +'/'+ child.getName());
						file.delete();
						System.out.println(fileLoc +'/'+ child.getName()+" is deleted from directory");
					}			
				}
			}
			//remove folder if empty
			directoryListing = dir.listFiles();
			if(directoryListing.length == 0){
				dir.delete();
				System.out.println("Deleting empty folder");
			}
			//recursion
			directoryListing = dir.listFiles();
			if(directoryListing!=null){
				for(File child :directoryListing){
					String extension = child.getName().substring(child.getName().lastIndexOf('.')+1,child.getName().length());
					if(vbaList.contains(extension)){
					//loop
					recursiveExtract(child);	
					
					}
					
				}
			}		
		}
	
	public static void recursiveExtract(File initFile) throws Exception{
		ArrayList<String> vbaList = new ArrayList<String>(Arrays.asList("doc","docx","rtf"));
		ArrayList<String> imageList = new ArrayList<String>(Arrays.asList("emf","wmf","jpg","jpeg","png"));
		String extension = initFile.getName().substring(initFile.getName().lastIndexOf('.')+1,initFile.getName().length());
		String filePath = initFile.getPath().substring(0,initFile.getPath().lastIndexOf("\\")+1);
		String fileName = initFile.getName().substring(0,initFile.getName().lastIndexOf('.'));
		String fileLoc = filePath+fileName;
		String input = initFile.getPath();
		//create an empty file to store modified file and embedded files
		//fileName as the folder name
		File dir = new File(fileLoc);	
		dir.mkdir();
		
		//extract all embedded doc in initFile
		try{
	        String[] arguments = new String[]{"-z", "--extract-dir="+fileLoc, input};
	        System.out.println("Using TIKA CLI to dedect embedded Files. Target Directory: "+ fileLoc);
	        TikaCLI.main(arguments);
	        
	    }
	    catch(Exception e){
	    	System.out.println("Exception in saveEmbedds, during search in File: " + input + "\r\nDetails: " + e);
	    }
		
		//get the imageName
		ArrayList<String> imageName = new ArrayList<String>(); 
		File [] directoryListing = dir.listFiles();
		Arrays.sort(directoryListing, LastModifiedFileComparator.LASTMODIFIED_COMPARATOR);	
		if(directoryListing!=null){
			for(File child :directoryListing){
				//get doc to rtf
				String childExtension = child.getName().substring(child.getName().lastIndexOf('.')+1,child.getName().length());
				//before delete useless file,
				//extract the name and rename file
				if(childExtension.equals("emf")){
					Parser parser = new AutoDetectParser();
					BodyContentHandler handler = new BodyContentHandler();
					Metadata metadata = new Metadata();   //empty metadata object 
					FileInputStream inputstream = new FileInputStream(child);
					ParseContext context = new ParseContext();
					parser.parse(inputstream, handler, metadata, context);

					// now this metadata object contains the extracted metadata of the given file.
					String childHandler = handler.toString();
					String childName = childHandler.substring(0,childHandler.lastIndexOf('.'));
					imageName.add(childName.trim().replaceAll("\\s","_"));
					inputstream.close();
				}							
			}
		}
		
		//delete image
		if(directoryListing!=null){
			for(File child :directoryListing){
				String childExtension = child.getName().substring(child.getName().lastIndexOf('.')+1,child.getName().length());
				if(imageList.contains(childExtension)){
					File file = new File(fileLoc +'/'+ child.getName());
					boolean isDelete = file.delete();
					System.out.println(fileLoc +'/'+ child.getName()+" is deleted from directory");
					System.out.println(isDelete);
				}
			}
		}
		for(String ss:imageName){
			System.out.println(ss);
		}
		
		//rename file
		//the file order of second layer
		directoryListing = dir.listFiles();
		Arrays.sort(directoryListing, LastModifiedFileComparator.LASTMODIFIED_COMPARATOR);		
		if(directoryListing!=null){
			int childIndex=0;
			for(File child :directoryListing){
				String childExtension = child.getName().substring(child.getName().lastIndexOf('.')+1,child.getName().length());
				if(!imageList.contains(childExtension)){
				System.out.println("File index: "+childIndex);
				File newChild = newFile(child,imageName.get(childIndex),childExtension);		
				System.out.println(child.getName()+" is renamed to "+ newChild.getName());
				boolean isSuccess = child.renameTo(newChild);
				System.out.println(isSuccess);
				childIndex++;
				}
			}
		}
		
		//delete non-whiteList doc
		directoryListing = dir.listFiles();
		if(directoryListing!=null){
			for(File child :directoryListing){
				String childExtension = child.getName().substring(child.getName().lastIndexOf('.')+1,child.getName().length());
				if(!whiteList.contains(childExtension)){
					File file = new File(fileLoc +'/'+ child.getName());
					file.delete();
					System.out.println(fileLoc +'/'+ child.getName()+" is deleted from directory");
				}			
			}
		}
		
		//data Cleansing of initFile
				if(extension.equals("docx")){
					
					mainHandler(initFile.getPath());
					
				}
				
				if(extension.equals("doc")){
					String newChildName = fileLoc+".docx";
					Document doc = new Document(initFile.getPath());		
					OoxmlSaveOptions saveOptions= new OoxmlSaveOptions();
					saveOptions.setSaveFormat(SaveFormat.DOCX);
					doc.save(newChildName, saveOptions);
					File newFile = new File(newChildName);
					
					mainHandler(newFile.getPath());
					
					initFile.delete();
					System.out.println(newChildName);
					
				}
			
		
		//remove folder if empty
		directoryListing = dir.listFiles();
		if(directoryListing.length == 0){
			dir.delete();
			System.out.println("Deleting empty folder");
		}
		
		//recursion
		directoryListing = dir.listFiles();
		if(directoryListing!=null){
			for(File child :directoryListing){
				String childExtension = child.getName().substring(child.getName().lastIndexOf('.')+1,child.getName().length());
				if(vbaList.contains(childExtension)){
				//loop
				recursiveExtract(child);	
				
				}
				
			}
		}		
	}

}

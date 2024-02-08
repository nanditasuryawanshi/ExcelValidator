package com.demo.ExcelProject;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HighlightChanges {

	private static final String FILE_PATH = "C:/Users/Dell/Desktop/data/MMS.xlsx";
	private static final String TARGET_SHEET_1 = "Clinic Master"; // Change to the desired sheet name
	private static final String TARGET_SHEET_2 = "Doctor Master";
	private static final String TARGET_SHEET_3 = "Package Master";
	private static final String TARGET_SHEET_4 = "Provider Master";
	private static final String TARGET_SHEET_5 = "Scheme Master";
	//similarly add for doctor etc ... private static final String TARGET_SHEET_NAME = "Clinic"; // Change to the desired sheet name 
	private static CellStyle errorStyle; // Declare a static errorStyle variable
    
	public static void main(String[] args) 
	{
		try 
		{
			FileInputStream file = new FileInputStream(new File(FILE_PATH));
			Workbook workbook = new XSSFWorkbook(file);
			DataFormatter dataFormatter = new DataFormatter();
			Iterator<Sheet> sheets = workbook.sheetIterator();
			while (sheets.hasNext())
			{
				Sheet sh = sheets.next();

				// Check if the sheet name is the target sheet 
				if (sh.getSheetName().equalsIgnoreCase(TARGET_SHEET_1)) 
				{
					System.out.println("Sheet name is " + sh.getSheetName());
					System.out.println("---------");
					Iterator<Row> iterator = sh.iterator();
					while (iterator.hasNext()) 
					{
						Row row = iterator.next();
						Iterator<Cell> cellIterator = row.iterator();
						while (cellIterator.hasNext()) 
						{
							Cell cell = cellIterator.next();
							String cellValue = dataFormatter.formatCellValue(cell);
							
							// Update cell values based on conditions--Mandatory check 
							if (cell.getColumnIndex() == 0 && (cellValue.equals("") || cellValue == null)) 
							{
								// Update cell with error message 
								//cell.setCellValue("Clinic Name-Eng is mandatory");
								setCellErrorStyle(cell);
							}
							// Update cell values based on conditions--Mandatory check 
							else if (cell.getColumnIndex() == 3 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
							
							// Check if column 3 has the value "Clinic/Clinic Only" --Address check 7,8,11,12, 13,16,17 
							Cell cellInColumn3 = row.getCell(3);
							String valueInColumn3 = dataFormatter.formatCellValue(cellInColumn3);
							if (valueInColumn3.contains("Clinic") || valueInColumn3.contains("Clinic Only"))
							{
									// Check if column 8 cell is null Address check 
								Cell cellInColumn8 = row.getCell(8);
								String valueInColumn8 = dataFormatter.formatCellValue(cellInColumn8);
								if (valueInColumn8.equals("") || valueInColumn8 == null) 
								{
									// Update cell with error message 
									//cellInColumn8.setCellValue("Column 8 is mandatory");
									setCellErrorStyle(cellInColumn8);
								}
								// Check if column 7 cell is null Address check 
								Cell cellInColumn7 = row.getCell(7);
								String valueInColumn7 = dataFormatter.formatCellValue(cellInColumn7);
								if (valueInColumn7.equals("") || valueInColumn7 == null) 
								{
									// Update cell with error message 
									//cellInColumn7.setCellValue("Column 7 is mandatory");
									setCellErrorStyle(cellInColumn7);
								}
									// Check if column 11 cell is null Address check 
								Cell cellInColumn11 = row.getCell(11);
								String valueInColumn11 = dataFormatter.formatCellValue(cellInColumn11);
								if (valueInColumn11.equals("") || valueInColumn11 == null) 
								{
									// Update cell with error message 
									//cellInColumn11.setCellValue("Column 11 is mandatory");
									setCellErrorStyle(cellInColumn11);
								}
								// Check if column 12 cell is null Address check 
								Cell cellInColumn12 = row.getCell(12);
								String valueInColumn12 = dataFormatter.formatCellValue(cellInColumn12);
								if (valueInColumn12.equals("") || valueInColumn12 == null) {
									// Update cell with error message 
									//cellInColumn12.setCellValue("Column 12 is mandatory");
									setCellErrorStyle(cellInColumn12);
								}
								// Check if column 13 cell is null Address check 
								Cell cellInColumn13 = row.getCell(13);
								String valueInColumn13 = dataFormatter.formatCellValue(cellInColumn13);
								if (valueInColumn13.equals("") || valueInColumn13 == null) {
									// Update cell with error message 
									//cellInColumn13.setCellValue("Column 13 is mandatory");
									setCellErrorStyle(cellInColumn13);
								}
								// Check if column 16 cell is null Address check 
								Cell cellInColumn16 = row.getCell(16);
								String valueInColumn16 = dataFormatter.formatCellValue(cellInColumn16);
								if (valueInColumn16.equals("") || valueInColumn16 == null) {
									// Update cell with error message 
									//cellInColumn16.setCellValue("Column 16 is mandatory");
									setCellErrorStyle(cellInColumn16);
								}
								// Check if column 17 cell is null Address check 
								Cell cellInColumn17 = row.getCell(17);
								String valueInColumn17 = dataFormatter.formatCellValue(cellInColumn17);
								if (valueInColumn17.equals("") || valueInColumn17 == null) {
									// Update cell with error message 
									//cellInColumn17.setCellValue("Column 17 is mandatory");
									setCellErrorStyle(cellInColumn17);
								}
							}
							// Check if column 3 has the values --contact mandatory,24,25,26 
							if (valueInColumn3.contains("Clinic") || valueInColumn3.contains("Clinic Only")|| valueInColumn3.contains("Surgical centre")|| valueInColumn3.contains("Hospital"))
							{
								// Check if column cell is null 
								Cell cellInColumn24 = row.getCell(24);
								String valueInColumn24 = dataFormatter.formatCellValue(cellInColumn24);
								if (valueInColumn24.equals("") || valueInColumn24 == null) 
								{
									// Update cell with error message 
									//cellInColumn24.setCellValue("contact mandatory");
									setCellErrorStyle(cellInColumn24);
								}
								// Check if column cell is null 
								Cell cellInColumn25 = row.getCell(25);
								String valueInColumn25 = dataFormatter.formatCellValue(cellInColumn25);
								if (valueInColumn25.equals("") || valueInColumn25 == null) 
								{
									// Update cell with error message 
									//cellInColumn25.setCellValue("contact mandatory");
									setCellErrorStyle(cellInColumn25);
								}
								// Check if column cell is null 
								Cell cellInColumn26 = row.getCell(26);
								String valueInColumn26 = dataFormatter.formatCellValue(cellInColumn26);
								if (valueInColumn26.equals("") || valueInColumn26 == null) {
									// Update cell with error message 
									//cellInColumn26.setCellValue("contact mandatory");
									setCellErrorStyle(cellInColumn26);
								}
								
							}
							
							// Check if column 3 has the value "Clinic Only" --Service Category mandatory,30
							if (valueInColumn3.contains("Clinic only")) 
							{
							    // Check if column cell is null
							    Cell cellInColumn30 = row.getCell(30);						
							    String valueInColumn30 = dataFormatter.formatCellValue(cellInColumn30);
							    if (valueInColumn30.equals("") || valueInColumn30 == null) 
							    {
							            // Update cell with error message
							            // cellInColumn30.setCellValue("Clinic only mandatory");
							            setCellErrorStyle(cellInColumn30);
							     }
							            // clinic only and sp then days mandatory
							     // Iterate through columns 32 to 63 with a step of 4
							    	 boolean hasNonNullValue = false;
							    	 for (int i = 32; i <= 63; i += 4) 
							    	 {
							    		 Cell cellInColumn = row.getCell(i);
							    		 if (cellInColumn != null) 
							    		 {
							    			 String valueInColumn = dataFormatter.formatCellValue(cellInColumn);
							    			 if (!(valueInColumn.equals("") || valueInColumn == null))
							    			 {// If at least one cell is not null, set the flag to true and break the loop
							    				 hasNonNullValue = true;
							                     break;
							                 }
							              }
							    	 }
							    	 boolean appoinment=false;
							    	 for (int i = 35; i <= 63; i += 4) 
							    	 {
							    		 Cell cellInColumn = row.getCell(i);							    		
							    		 if (cellInColumn != null) 
							    		 {
							                        String valueInColumn = dataFormatter.formatCellValue(cellInColumn);
							                        if (valueInColumn.equals("Y"))
							                        {
							                            // If at least one cell is not null, set the flag to true and break the loop
							                            appoinment = true;
							                            break;
							                        }
							              }
							    	 }							    	 
							    	 // If all cells with an index divisible by 4 are null, set the error message
							    	 if (!hasNonNullValue && !appoinment) 
							    	 {
							    		 for (int i = 32; i <= 63; i += 4) 
							    		 {
							    			 Cell cellInColumn = row.getCell(i);
							    			 if (cellInColumn != null) 
							    			 { //cellInColumn.setCellValue("Days are mandatory");
							    				 setCellErrorStyle(cellInColumn);
							    			 }
							    	      }
							    		 for (int i = 35; i <= 63; i += 4) 
							    		 {
							    			 Cell cellInColumn = row.getCell(i);
							    			 if (cellInColumn != null) 
							    			 { //cellInColumn.setCellValue("appointment is mandatory");
							    				 setCellErrorStyle(cellInColumn);
							    			 }
							    	      }
							         }							    	 
							    
							}						    
							//--------------------------------------------------------------------------------------------------------
							Cell cellInColumn32 = row.getCell(32);
                            String valueInColumn32 = dataFormatter.formatCellValue(cellInColumn32);
                            
                            //new
                            //String[] daysArray = {"Mon", "Tue", "Wed", "Thur", "Fri", "Sat", "Sun", "Public Holiday"};

                    	    for (int i = 32; i <= 63; i += 4) 
                    	    {
                    	        Cell cellInColumn = row.getCell(i);
                    	        String valueInColumn = dataFormatter.formatCellValue(cellInColumn);
                    	      
                    	        if(valueInColumn.contains("Mon")||valueInColumn.contains("Tue")||valueInColumn.contains("Wed")||valueInColumn.contains("Thur")||valueInColumn.contains("Fri")||valueInColumn.contains("Sat")||valueInColumn.contains("Sun")||valueInColumn.contains("Public Holiday"))
                                { // Check the next column (i + 4)
                    	            Cell nextCell = row.getCell(i + 1);
                    	            String valueInNextColumn = dataFormatter.formatCellValue(nextCell);

                    	            // Check the column after the next one (i + 8)
                    	            Cell nextNextCell = row.getCell(i + 2);
                    	            String valueInNextNextColumn = dataFormatter.formatCellValue(nextNextCell);

                    	            // Check if the values in the next two columns are not null
                    	            if (valueInNextColumn.equals("") || valueInNextColumn == null || valueInNextNextColumn.equals("") || valueInNextNextColumn == null) 
                    	            {
                    	                setCellErrorStyle(nextCell);
                    	                setCellErrorStyle(nextNextCell);
                    	            }
                                }
                    	        else if(!(valueInColumn.equals("") || valueInColumn ==null || valueInColumn.contains("Servicing Days")))
                    	        {
                    	        	 // Update cell with error message
                                    //cellInColumn.setCellValue("Error in day");
                                    setCellErrorStyle(cellInColumn);                  	     	
                    	        }
                    	    }
                    	    
                    	    
                    	    
							//optional if Clinic/ Clinic Only/Surgical Centre/Hospital  

							if (valueInColumn3.contains("China Hospital") || valueInColumn3.contains("HK Public HOSP")|| valueInColumn3.contains("Macau Desiginated HOSP")|| valueInColumn3.contains("Others"))
							{
								// Check if column cell is null 
								Cell cellInColumn4 = row.getCell(4);
								String valueInColumn4 = dataFormatter.formatCellValue(cellInColumn4);
								if (valueInColumn4.equals("") || valueInColumn4 == null) {
									// Update cell with error message 
									//cellInColumn4.setCellValue("mandatory");
									setCellErrorStyle(cellInColumn4);
								}
							}
							
							
							// Check if column 30 has "Service Category = SP" --Specility Category mandatory,31 
							Cell cellInColumn30 = row.getCell(30);
							String valueInColumn30 = dataFormatter.formatCellValue(cellInColumn30);
							if (valueInColumn30.contains("SP"))
							{
								// Check if column cell is null 
								Cell cellInColumn31 = row.getCell(31);
								String valueInColumn31 = dataFormatter.formatCellValue(cellInColumn31);
								if (valueInColumn31.equals("") || valueInColumn31 == null) {
									// Update cell with error message 
									//cellInColumn31.setCellValue("SP --Specility Category mandatory");
									setCellErrorStyle(cellInColumn31);
								}
							}
							
							 String[] invalidCharacters = {"*", ";","@"};

	                         // Assuming 'cellValue' is the value you are checking
	                         for (String invalidChar : invalidCharacters) {
	                             if (cellValue.contains(invalidChar)) {
	                                 // Update cell with error message
	                                 //cell.setCellValue(cellValue + " INCORRECT VALUE");
	                                 setCellErrorStyle(cell);
	                                 break; // exit the loop once a match is found
	                             }
	                         }
	                         
	                      // Check if "China Hospital" --traditional name check 
								if (valueInColumn3.contains("China Hospital"))
								{
									// Check if column cell is null 
									Cell cellInColumn1 = row.getCell(1);
									String valueInColumn1 = dataFormatter.formatCellValue(cellInColumn1);
									if (valueInColumn1.equals("") || valueInColumn1 == null) {
										// Update cell with error message 
										//cellInColumn1.setCellValue("mandatory because China Hospital");
										setCellErrorStyle(cellInColumn1);
									}
								}
	                         
	                         //check for contact phone code
	                         Cell cellInColumn25 = row.getCell(25);
							String valueInColumn25 = dataFormatter.formatCellValue(cellInColumn25);
							//added 6 feb
							int len= valueInColumn25.length();
							if (!(valueInColumn25.equals("") || valueInColumn25 == null)) 
							{
								
								if(!(valueInColumn25.contains("853")||valueInColumn25.contains("852")||valueInColumn25.contains("86")))
								{  // Update cell with error message
									// cellInColumn24.setCellValue("Invalid length (should be 8 digits)");
									setCellErrorStyle(cellInColumn25);
									
								}
							}
							
							//check for contact phone 8-digits
							Cell cellInColumn26 = row.getCell(26);
							String valueInColumn26 = dataFormatter.formatCellValue(cellInColumn26);
							
							if (!(valueInColumn26=="" || valueInColumn26 == null))
							{
							if(!(valueInColumn26.matches("\\d{8}")))
							{  // Update cell with error message
								// cellInColumn24.setCellValue("Invalid length (should be 8 digits)");
								setCellErrorStyle(cellInColumn26);
								
							}
							        
							   	
						}
							
						}
					}
				}
				//DOCTOR=================================================================================================
												
				else if (sh.getSheetName().equalsIgnoreCase(TARGET_SHEET_2))
				{
					// 	Doctor doctor=new Doctor(); 
					// 	doctor.display(); 
					System.out.println("Sheet name is " + sh.getSheetName());
					System.out.println("---------");
					Doctor d=new Doctor();
					d.display();
					
					
					

				
					Iterator<Row> iterator = sh.iterator();
					while (iterator.hasNext()) 
					{
						Row row = iterator.next();
						Iterator<Cell> cellIterator = row.iterator();
						while (cellIterator.hasNext()) 
						{
							Cell cell = cellIterator.next();
							String cellValue = dataFormatter.formatCellValue(cell);
							
							// Update cell values based on conditions--Mandatory check 
							if (cell.getColumnIndex() == 0 && (cellValue.equals("") || cellValue == null)) 
							{
								// Update cell with error message 
								//cell.setCellValue("Clinic Name-Eng is mandatory");
								setCellErrorStyle(cell);
							}
							// Update cell values based on conditions--Mandatory check 
							if (cell.getColumnIndex() == 1 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
							
							if (cell.getColumnIndex() == 4 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
							
							if (cell.getColumnIndex() == 5 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
							
							
							if (cell.getColumnIndex() == 7 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
							
							
							// Check if column 30 has "Service Category = SP" --Specility Category mandatory,31 
							Cell cellInColumn5 = row.getCell(5);
							String valueInColumn5 = dataFormatter.formatCellValue(cellInColumn5);
							if (valueInColumn5.contains("SP"))
							{
								// Check if column cell is null 
								Cell cellInColumn6 = row.getCell(6);
								String valueInColumn6 = dataFormatter.formatCellValue(cellInColumn6);
								if (valueInColumn6.equals("") || valueInColumn6 == null) {
									// Update cell with error message 
									//cellInColumn31.setCellValue("SP --Specility Category mandatory");
									setCellErrorStyle(cellInColumn6);
								}
							}
							
							 String[] invalidCharacters = {"*", ";","@"};

	                         // Assuming 'cellValue' is the value you are checking
	                         for (String invalidChar : invalidCharacters) {
	                             if (cellValue.contains(invalidChar)) {
	                                 // Update cell with error message
	                                 //cell.setCellValue(cellValue + " INCORRECT VALUE");
	                                 setCellErrorStyle(cell);
	                                 break; // exit the loop once a match is found
	                             }
	                         }
	                    
							
						}
					}
				
				}
				
				//PACKAGE=================================================================================================
				
				else if (sh.getSheetName().equalsIgnoreCase(TARGET_SHEET_3))
				{
					// 	Doctor doctor=new Doctor(); 
					// 	doctor.display(); 
					System.out.println("Sheet name is " + sh.getSheetName());
					System.out.println("---------");
					
					

				
					Iterator<Row> iterator = sh.iterator();
					while (iterator.hasNext()) 
					{
						Row row = iterator.next();
						Iterator<Cell> cellIterator = row.iterator();
						while (cellIterator.hasNext()) 
						{
							Cell cell = cellIterator.next();
							String cellValue = dataFormatter.formatCellValue(cell);
							
							// Update cell values based on conditions--Mandatory check 
							if (cell.getColumnIndex() == 0 && (cellValue.equals("") || cellValue == null)) 
							{
								// Update cell with error message 
								//cell.setCellValue("Clinic Name-Eng is mandatory");
								setCellErrorStyle(cell);
							}
							// Update cell values based on conditions--Mandatory check 
							if (cell.getColumnIndex() == 1 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
							
							if (cell.getColumnIndex() == 2 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
							
							if (cell.getColumnIndex() == 3 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
							
							
							if (cell.getColumnIndex() == 4 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
							
							if (cell.getColumnIndex() == 5 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
							if (cell.getColumnIndex() == 6 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
						
							
						
							
							 String[] invalidCharacters = {"*", ";","@"};

	                         // Assuming 'cellValue' is the value you are checking
	                         for (String invalidChar : invalidCharacters) {
	                             if (cellValue.contains(invalidChar)) {
	                                 // Update cell with error message
	                                 //cell.setCellValue(cellValue + " INCORRECT VALUE");
	                                 setCellErrorStyle(cell);
	                                 break; // exit the loop once a match is found
	                             }
	                         }
	                    
							
						}
					}
				
				}
				
				//PROVIDER=================================================================================================
				
				else if (sh.getSheetName().equalsIgnoreCase(TARGET_SHEET_4))
				{
					// 	Doctor doctor=new Doctor(); 
					// 	doctor.display(); 
					System.out.println("Sheet name is " + sh.getSheetName());
					System.out.println("---------");
					

				
					Iterator<Row> iterator = sh.iterator();
					while (iterator.hasNext()) 
					{
						Row row = iterator.next();
						Iterator<Cell> cellIterator = row.iterator();
						while (cellIterator.hasNext()) 
						{
							Cell cell = cellIterator.next();
							String cellValue = dataFormatter.formatCellValue(cell);
							
							// Update cell values based on conditions--Mandatory check 
							if (cell.getColumnIndex() == 0 && (cellValue.equals("") || cellValue == null)) 
							{
								// Update cell with error message 
								//cell.setCellValue("Clinic Name-Eng is mandatory");
								setCellErrorStyle(cell);
							}
							// Update cell values based on conditions--Mandatory check 
							if (cell.getColumnIndex() == 1 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
							
							
							if (cell.getColumnIndex() == 4 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
							
							if (cell.getColumnIndex() == 5 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
							
							 String[] invalidCharacters = {"*", ";","@"};

	                         // Assuming 'cellValue' is the value you are checking
	                         for (String invalidChar : invalidCharacters) {
	                             if (cellValue.contains(invalidChar)) {
	                                 // Update cell with error message
	                                 //cell.setCellValue(cellValue + " INCORRECT VALUE");
	                                 setCellErrorStyle(cell);
	                                 break; // exit the loop once a match is found
	                             }
	                         }
	                    
							
						}
					}
				
				}				
				//SCHEME=====================================================				
				else if (sh.getSheetName().equalsIgnoreCase(TARGET_SHEET_5))
				{
					// 	Doctor doctor=new Doctor(); 
					// 	doctor.display(); 
					System.out.println("Sheet name is " + sh.getSheetName());
					System.out.println("---------");
					
					Iterator<Row> iterator = sh.iterator();
					while (iterator.hasNext()) 
					{
						Row row = iterator.next();
						Iterator<Cell> cellIterator = row.iterator();
						while (cellIterator.hasNext()) 
						{
							Cell cell = cellIterator.next();
							String cellValue = dataFormatter.formatCellValue(cell);
						
							if (cell.getColumnIndex() == 1 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
							
							
							if (cell.getColumnIndex() == 2 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
							
							if (cell.getColumnIndex() == 3 && (cellValue.equals("") || cellValue == null)) {
								// Update cell with error message 
								//cell.setCellValue("Clinic Type is mandatory");
								setCellErrorStyle(cell);
							}
							
							 boolean hasNonNullValue = false;
					    	 for (int i = 4; i <= 15; i ++) 
					    	 {
					    		 Cell cellInColumn = row.getCell(i);
					    		 if (cellInColumn != null) 
					    		 {
					    			 String valueInColumn = dataFormatter.formatCellValue(cellInColumn);
					    			 if (!(valueInColumn.equals("") || valueInColumn == null))
					    			 {// If at least one cell is not null, set the flag to true and break the loop
					    				 hasNonNullValue = true;
					                     break;
					                 }
					              }
					    	 }
					    	 if (!hasNonNullValue) 
					    	 {
					    		 for (int i = 4; i <= 15; i ++) 
					    		 {
					    			 Cell cellInColumn = row.getCell(i);
					    			 if (cellInColumn != null) 
					    			 { //cellInColumn.setCellValue("Days are mandatory");
					    				 setCellErrorStyle(cellInColumn);
					    			 }
					    		 }
					    	    
					         }		
					    	 
							
							 String[] invalidCharacters = {"*", ";","@"};

	                         // Assuming 'cellValue' is the value you are checking
	                         for (String invalidChar : invalidCharacters) {
	                             if (cellValue.contains(invalidChar)) {
	                                 // Update cell with error message
	                                 //cell.setCellValue(cellValue + " INCORRECT VALUE");
	                                 setCellErrorStyle(cell);
	                                 break; // exit the loop once a match is found
	                             }
	                         }
	                    
							
						}
					}
				
				}
				//==============================================================================
				else 
				{
					System.out.println("Skipping sheet: " + sh.getSheetName());
				}
			}
			file.close(); // Close the FileInputStream
			// Save the changes to the Excel file 
			try (FileOutputStream outputStream = new FileOutputStream(("C:/Users/Dell/Desktop/data/MMS2.xlsx"))) {
				workbook.write(outputStream);
				workbook.close();
			}
		} 
		
		
		catch (Exception e) 
		{
			e.printStackTrace();
		}
	}
	private static void setCellErrorStyle(Cell cell) {
		Workbook workbook = cell.getSheet().getWorkbook();
		if (errorStyle == null) {
			// Create the style only if it's not already created 
			errorStyle = workbook.createCellStyle();
			errorStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
			errorStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		}
		// Set the error style to the cell 
		cell.setCellStyle(errorStyle);
	}
}

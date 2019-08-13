public String getCellData(String sheetName,String colName, int rowNum)
{
try{
}
catch{
}public String getCellData(String sheetName,String colName, int rowNum)
	{
	try{
		int colNum= -1;
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(0);
		for (int i=0;i<row.getLastCellNum();i++) {}
		if(row.getCell(i).getStringCellValue().trim().equals(colName))
			colName=i;
	}
	row= sheet.getRow(rowNum-1);
	cell = row.getCell(colNum);
	return cell.getStringCellValue();
	
	catch(Exception e){
	}
}

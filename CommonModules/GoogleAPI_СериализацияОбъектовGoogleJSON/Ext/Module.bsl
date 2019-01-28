#Область СериализацияОбъектовGoogleJSON

////////////////////////////////////////////////////////////////////////////////
// Функции заполнения объектов Google Spreadsheet
//

Процедура ЗаполнитьОбъектПоУмолчаниюGSs(ОбъектGSs, ФабрикаGSs = Неопределено, ДопПараметры = Неопределено) Экспорт 
	
	Если ОбъектGSs.Тип() = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "Spreadsheet") Тогда
		ЗаполнитьПоУмолчаниюSpreadsheet(ОбъектGSs, ФабрикаGSs, ДопПараметры);
		
	ИначеЕсли ОбъектGSs.Тип() = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "SpreadsheetProperties") Тогда
		ЗаполнитьПоУмолчаниюSpreadsheetProperties(ОбъектGSs, ФабрикаGSs);
		
	ИначеЕсли ОбъектGSs.Тип() = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "Sheet") Тогда
		ЗаполнитьПоУмолчаниюSheet(ОбъектGSs, ФабрикаGSs);
	
	ИначеЕсли ОбъектGSs.Тип() = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "SheetProperties") Тогда
		ЗаполнитьПоУмолчаниюSheetProperties(ОбъектGSs, ФабрикаGSs);
	
	ИначеЕсли ОбъектGSs.Тип() = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "CellFormat") Тогда
		ЗаполнитьПоУмолчаниюCellFormat(ОбъектGSs, ФабрикаGSs);
		
	ИначеЕсли ОбъектGSs.Тип() = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "Color") Тогда
		ЗаполнитьПоУмолчаниюColor(ОбъектGSs, ФабрикаGSs);
		
	ИначеЕсли ОбъектGSs.Тип() = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "Padding") Тогда
		ЗаполнитьПоУмолчаниюPadding(ОбъектGSs, ФабрикаGSs);
		
	ИначеЕсли ОбъектGSs.Тип() = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "TextFormat") Тогда
		ЗаполнитьПоУмолчаниюTextFormat(ОбъектGSs, ФабрикаGSs);
		
	ИначеЕсли ОбъектGSs.Тип() = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "GridProperties") Тогда
		ЗаполнитьПоУмолчаниюGridProperties(ОбъектGSs, ФабрикаGSs);
		
	ИначеЕсли ОбъектGSs.Тип() = ФабрикаGSs.Тип("http://www.GoogleAPI_GoogleDrive.org", "GoogleDriveFile") Тогда
		ЗаполнитьПоУмолчаниюParentsFile(ОбъектGSs, ФабрикаGSs, ДопПараметры);
		
	КонецЕсли; 

КонецПроцедуры // ЗаполнитьGSОбъектПоУмолчанию()

Процедура ЗаполнитьПоУмолчаниюSpreadsheet(ОбъектGSs, ФабрикаGSs, ДопПараметры = Неопределено)

	КоличествоЛистов = 1;
	
	Если НЕ ДопПараметры = Неопределено Тогда
		КоличествоЛистов = ?(ДопПараметры.Свойство("КоличествоЛистов",КоличествоЛистов), КоличествоЛистов, 1);
	КонецЕсли; 
	
	ОбъектGSs.spreadsheetId = "";

	ТипSpreadsheetProperties = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "SpreadsheetProperties");
	ОбъектSpreadsheetProperties = ФабрикаGSs.Создать(ТипSpreadsheetProperties);
	ЗаполнитьОбъектПоУмолчаниюGSs(ОбъектSpreadsheetProperties, ФабрикаGSs);
	
	ОбъектGSs.properties = ОбъектSpreadsheetProperties;
	
	Для Сч = 1 По КоличествоЛистов Цикл
	
		ТипSheet = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "Sheet");
		ОбъектSheet = ФабрикаGSs.Создать(ТипSheet);
		ЗаполнитьОбъектПоУмолчаниюGSs(ОбъектSheet, ФабрикаGSs);
		
		ОбъектSheet.properties.sheetId = Сч-1;
		ОбъектSheet.properties.title =	"Sheet" + Строка(Сч-1);
		ОбъектSheet.properties.index =	Сч-1;
		
		ОбъектGSs.sheets.Добавить(ОбъектSheet);
	
	КонецЦикла; 
	
	ОбъектGSs.spreadsheetUrl = "";
	
КонецПроцедуры // ЗаполнитьПоУмолчаниюSpreadsheet()

Процедура ЗаполнитьПоУмолчаниюSheet(ОбъектGSs, ФабрикаGSs)

	ТипSheetProperties = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "SheetProperties");
	ОбъектSheetProperties = ФабрикаGSs.Создать(ТипSheetProperties);
	ЗаполнитьОбъектПоУмолчаниюGSs(ОбъектSheetProperties, ФабрикаGSs);
	
	ОбъектGSs.properties = ОбъектSheetProperties;
	
КонецПроцедуры // ЗаполнитьПоУмолчаниюSheet()

Процедура ЗаполнитьПоУмолчаниюSheetProperties(ОбъектGSs, ФабрикаGSs)

	ОбъектGSs.sheetId	 = 0;
	ОбъектGSs.title		 = "Sheet1";
	ОбъектGSs.index		 = 0;
	ОбъектGSs.sheetType	 = "GRID";
	
	ТипGridProperties = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "GridProperties");
	ОбъектGridProperties = ФабрикаGSs.Создать(ТипGridProperties);
	ЗаполнитьОбъектПоУмолчаниюGSs(ОбъектGridProperties, ФабрикаGSs);
	
	ОбъектGSs.gridProperties = ОбъектGridProperties;
	
КонецПроцедуры // ЗаполнитьПоУмолчаниюSheetProperties()

Процедура ЗаполнитьПоУмолчаниюSpreadsheetProperties(ОбъектGSs, ФабрикаGSs)

	ОбъектGSs.title		 = "Новая таблица";
	ОбъектGSs.locale	 = "ru_RU";
	ОбъектGSs.autoRecalc = "ON_CHANGE";
	ОбъектGSs.timeZone	 = "Etc/GMT";

	ТипCellFormat = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "CellFormat");
	ОбъектCellFormat = ФабрикаGSs.Создать(ТипCellFormat);
	ЗаполнитьОбъектПоУмолчаниюGSs(ОбъектCellFormat, ФабрикаGSs);
	
	ОбъектGSs.defaultFormat = ОбъектCellFormat;
	
КонецПроцедуры // ЗаполнитьПоУмолчаниюSpreadsheetProperties()

Процедура ЗаполнитьПоУмолчаниюCellFormat(ОбъектGSs, ФабрикаGSs)

	ТипColor = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "Color");
	ОбъектColor = ФабрикаGSs.Создать(ТипColor);
	ЗаполнитьОбъектПоУмолчаниюGSs(ОбъектColor, ФабрикаGSs);
	
	ОбъектGSs.backgroundColor = ОбъектColor;

	ТипPadding = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "Padding");
	ОбъектPadding = ФабрикаGSs.Создать(ТипPadding);
	ЗаполнитьОбъектПоУмолчаниюGSs(ОбъектPadding, ФабрикаGSs);
	
	ОбъектGSs.padding			 = ОбъектPadding;
	ОбъектGSs.verticalAlignment	 = "BOTTOM";
	ОбъектGSs.wrapStrategy		 = "OVERFLOW_CELL";
	
	ТипTextFormat = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "TextFormat");
	ОбъектTextFormat = ФабрикаGSs.Создать(ТипTextFormat);
	ЗаполнитьОбъектПоУмолчаниюGSs(ОбъектTextFormat, ФабрикаGSs);
	
	ОбъектGSs.textFormat = ОбъектTextFormat;
	
КонецПроцедуры // ЗаполнитьПоУмолчаниюCellFormat()

Процедура ЗаполнитьПоУмолчаниюColor(ОбъектGSs, ФабрикаGSs)

	ОбъектGSs.red	 = 1;
	ОбъектGSs.green	 = 1;
	ОбъектGSs.blue	 = 1;

КонецПроцедуры // ЗаполнитьПоУмолчаниюColor()

Процедура ЗаполнитьПоУмолчаниюPadding(ОбъектGSs, ФабрикаGSs)

	ОбъектGSs.top	 = 2;
	ОбъектGSs.right	 = 3;
	ОбъектGSs.bottom = 2;
	ОбъектGSs.left	 = 3;

КонецПроцедуры // ЗаполнитьПоУмолчаниюPadding()

Процедура ЗаполнитьПоУмолчаниюTextFormat(ОбъектGSs, ФабрикаGSs)

	ОбъектGSs.foregroundColor	 = Неопределено;
	ОбъектGSs.fontFamily		 = "arial,sans,sans-serif";
	ОбъектGSs.fontSize			 = 10;
	ОбъектGSs.bold				 = Ложь;
	ОбъектGSs.italic			 = Ложь;
	ОбъектGSs.strikethrough		 = Ложь;
	ОбъектGSs.underline			 = Ложь;

КонецПроцедуры // ЗаполнитьПоУмолчаниюTextFormat()

Процедура ЗаполнитьПоУмолчаниюGridProperties(ОбъектGSs, ФабрикаGSs)

	ОбъектGSs.rowCount		 = 1000;
	ОбъектGSs.columnCount	 = 26;

КонецПроцедуры // ЗаполнитьПоУмолчаниюGridProperties()

Процедура ЗаполнитьОбъектПоУмолчаниюGSsUpdate(ОбъектGSs, ФабрикаGSs = Неопределено, ДопПараметры = Неопределено)

	Если ФабрикаGSs = Неопределено Тогда
		//todo: определить фабрику
	КонецЕсли; 
	
	Если ОбъектGSs.Тип() = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "BatchUpdate") Тогда
		ЗаполнитьПоУмолчаниюBatchUpdate(ОбъектGSs, ФабрикаGSs);		
	КонецЕсли; 

КонецПроцедуры // ЗаполнитьОбъектПоУмолчаниюGSsUpdate()

Процедура ЗаполнитьПоУмолчаниюBatchUpdate(ОбъектGSs, ФабрикаGSs)

	ОбъектGSs.title		 = "Новая таблица";
	ОбъектGSs.locale	 = "ru_RU";
	ОбъектGSs.autoRecalc = "ON_CHANGE";
	ОбъектGSs.timeZone	 = "Etc/GMT";

	ТипCellFormat = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "CellFormat");
	ОбъектCellFormat = ФабрикаGSs.Создать(ТипCellFormat);
	ЗаполнитьОбъектПоУмолчаниюGSs(ОбъектCellFormat, ФабрикаGSs);
	
	ОбъектGSs.defaultFormat = ОбъектCellFormat;
	
КонецПроцедуры // ЗаполнитьПоУмолчаниюSpreadsheetProperties()

Функция ПолучитьОбъектЗапросаУстановкаШириныКолонки(ФабрикаGSs, ЛистИД, ПерваяКолонка, ПоследняяКолонка, ШиринаКолонки) Экспорт 

	ТипBatchRequestUpdateDimensionProperties = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "BatchRequestUpdateDimensionProperties");
	ТипUpdateDimensionProperties = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "UpdateDimensionProperties");
	ТипDimensionProperties		 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "DimensionProperties");
	ТипDimensionRange			 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "DimensionRange");
	
	ОбъектBatchRequestUpdateDimensionProperties = ФабрикаGSs.Создать(ТипBatchRequestUpdateDimensionProperties);
	
	ОбъектUpdateDimensionProperties = ФабрикаGSs.Создать(ТипUpdateDimensionProperties);
	
	ОбъектDimensionRange = ФабрикаGSs.Создать(ТипDimensionRange);
	ОбъектDimensionRange.sheetId	 = ЛистИД;
	ОбъектDimensionRange.dimension	 = "COLUMNS";
	ОбъектDimensionRange.startIndex	 = ПерваяКолонка - 1;
	ОбъектDimensionRange.endIndex	 = ПоследняяКолонка;
	ОбъектUpdateDimensionProperties.range = ОбъектDimensionRange;
	
	ОбъектDimensionProperties = ФабрикаGSs.Создать(ТипDimensionProperties);
	ОбъектDimensionProperties.pixelSize = ШиринаКолонки;
	ОбъектUpdateDimensionProperties.properties = ОбъектDimensionProperties;
	
	ОбъектUpdateDimensionProperties.fields = "pixelSize";
	
	ОбъектBatchRequestUpdateDimensionProperties.updateDimensionProperties = ОбъектUpdateDimensionProperties;	

	Возврат ОбъектBatchRequestUpdateDimensionProperties;
	
КонецФункции // ПолучитьЗапросУстановкаШириныКолонки()

Функция ПолучитьОбъектBatchRequestRepeatCellВыравнивание(ФабрикаGSs, ЛистИД, ПерваяКолонка = Неопределено, ПоследняяКолонка = Неопределено, 
		ПерваяСтрока = Неопределено, ПоследняяСтрока = Неопределено, ГоризонтальноеПоложение = Неопределено, ВертикальноеПоложение = Неопределено) Экспорт

	ТипBatchRequestRepeatCell	 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "BatchRequestRepeatCell");
	ТипRepeatCell				 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "RepeatCell");
	ТипGridRange				 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "GridRange");
	ТипCellData					 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "CellData");
	ТипCellFormat				 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "CellFormat");
	
	ОбъектBatchRequestRepeatCell = ФабрикаGSs.Создать(ТипBatchRequestRepeatCell);
	
	ОбъектRepeatCell = ФабрикаGSs.Создать(ТипRepeatCell);
	
	ОбъектGridRange = ФабрикаGSs.Создать(ТипGridRange);
	
	ОбъектGridRange.sheetId			 = ЛистИД;
	
	Если НЕ ПерваяКолонка = Неопределено Тогда
		ОбъектGridRange.startColumnIndex = ПерваяКолонка - 1;
	КонецЕсли; 
	Если НЕ ПоследняяКолонка = Неопределено Тогда
		ОбъектGridRange.endColumnIndex = ПоследняяКолонка;
	КонецЕсли;
	Если НЕ ПерваяСтрока = Неопределено Тогда
		ОбъектGridRange.startRowIndex = ПерваяСтрока - 1;
	КонецЕсли;
	Если НЕ ПоследняяСтрока = Неопределено Тогда
		ОбъектGridRange.endRowIndex = ПоследняяСтрока;
	КонецЕсли;
	
	ОбъектRepeatCell.range = ОбъектGridRange;

	ОбъектCellData = ФабрикаGSs.Создать(ТипCellData);
	
	ОбъектCellFormat = ФабрикаGSs.Создать(ТипCellFormat);
	
	ОбъектCellFormat.horizontalAlignment = ?(НЕ ЗначениеЗаполнено(ГоризонтальноеПоложение), "LEFT", ГоризонтальноеПоложение);
	ОбъектCellFormat.verticalAlignment	 = ?(НЕ ЗначениеЗаполнено(ВертикальноеПоложение), "TOP", ВертикальноеПоложение);
	ОбъектCellFormat.wrapStrategy		 = "WRAP";
	
	ОбъектCellData.userEnteredFormat = ОбъектCellFormat;
	
	ОбъектRepeatCell.fields = "userEnteredFormat(horizontalAlignment,verticalAlignment,wrapStrategy)";
	
	ОбъектRepeatCell.cell = ОбъектCellData;
	
	ОбъектBatchRequestRepeatCell.repeatCell = ОбъектRepeatCell;
	
	Возврат ОбъектBatchRequestRepeatCell;
	
КонецФункции // ПолучитьОбъектBatchRequestRepeatCell()

Функция ПолучитьОбъектBatchRequestRepeatCellФонЯчейки(ФабрикаGSs, ЛистИД, ПерваяКолонка, ПоследняяКолонка, ПерваяСтрока, ПоследняяСтрока, Цвет) Экспорт

	ТипBatchRequestRepeatCell	 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "BatchRequestRepeatCell");
	ТипRepeatCell				 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "RepeatCell");
	ТипGridRange				 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "GridRange");
	ТипCellData					 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "CellData");
	ТипCellFormat				 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "CellFormat");
	ТипColor					 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "Color");
	
	ОбъектBatchRequestRepeatCell = ФабрикаGSs.Создать(ТипBatchRequestRepeatCell);
	
	ОбъектRepeatCell = ФабрикаGSs.Создать(ТипRepeatCell);
	
	ОбъектGridRange = ФабрикаGSs.Создать(ТипGridRange);
	
	ОбъектGridRange.sheetId			 = ЛистИД;
	ОбъектGridRange.startColumnIndex = ПерваяКолонка - 1;
	ОбъектGridRange.endColumnIndex	 = ПоследняяКолонка;
	ОбъектGridRange.startRowIndex	 = ПерваяСтрока - 1;
	ОбъектGridRange.endRowIndex		 = ПоследняяСтрока;
	
	ОбъектRepeatCell.range = ОбъектGridRange;

	ОбъектCellData = ФабрикаGSs.Создать(ТипCellData);
	
	ОбъектCellFormat = ФабрикаGSs.Создать(ТипCellFormat);
	
	ОбъектColor = ФабрикаGSs.Создать(ТипColor);
	
	ОбъектColor.red	 = Цвет.Красный;
	ОбъектColor.green = Цвет.Зеленый;
	ОбъектColor.blue = Цвет.Синий;
	
	ОбъектCellFormat.backgroundColor = ОбъектColor;
	
	ОбъектCellData.userEnteredFormat = ОбъектCellFormat;
	
	ОбъектRepeatCell.fields = "userEnteredFormat(backgroundColor)";
	
	ОбъектRepeatCell.cell = ОбъектCellData;
	
	ОбъектBatchRequestRepeatCell.repeatCell = ОбъектRepeatCell;
	
	Возврат ОбъектBatchRequestRepeatCell;
	
КонецФункции // ПолучитьОбъектBatchRequestRepeatCell()

Функция ПолучитьОбъектBatchRequestRepeatCellФорматЯчеек(ФабрикаGSs, ЛистИД, ПерваяКолонка, ПоследняяКолонка, ПерваяСтрока, ПоследняяСтрока, ФорматТип, ФорматШаблон) Экспорт

	ТипBatchRequestRepeatCell	 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "BatchRequestRepeatCell");
	ТипRepeatCell				 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "RepeatCell");
	ТипGridRange				 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "GridRange");
	ТипCellData					 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "CellData");
	ТипCellFormat				 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "CellFormat");
	ТипNumberFormat					 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "NumberFormat");
	
	ОбъектBatchRequestRepeatCell = ФабрикаGSs.Создать(ТипBatchRequestRepeatCell);
	
	ОбъектRepeatCell = ФабрикаGSs.Создать(ТипRepeatCell);
	
	ОбъектGridRange = ФабрикаGSs.Создать(ТипGridRange);
	
	ОбъектGridRange.sheetId			 = ЛистИД;
	ОбъектGridRange.startColumnIndex = ПерваяКолонка - 1;
	ОбъектGridRange.endColumnIndex	 = ПоследняяКолонка;
	ОбъектGridRange.startRowIndex	 = ПерваяСтрока - 1;
	ОбъектGridRange.endRowIndex		 = ПоследняяСтрока;
	
	ОбъектRepeatCell.range = ОбъектGridRange;

	ОбъектCellData = ФабрикаGSs.Создать(ТипCellData);
	
	ОбъектCellFormat = ФабрикаGSs.Создать(ТипCellFormat);
	
	ОбъектNumberFormat = ФабрикаGSs.Создать(ТипNumberFormat);
	
	ОбъектNumberFormat.type		 = ФорматТип;
	ОбъектNumberFormat.pattern	 = ФорматШаблон;
	
	ОбъектCellFormat.numberFormat = ОбъектNumberFormat;
	
	ОбъектCellData.userEnteredFormat = ОбъектCellFormat;
	
	ОбъектRepeatCell.fields = "userEnteredFormat(numberFormat)";
	
	ОбъектRepeatCell.cell = ОбъектCellData;
	
	ОбъектBatchRequestRepeatCell.repeatCell = ОбъектRepeatCell;
	
	Возврат ОбъектBatchRequestRepeatCell;

КонецФункции // ПолучитьОбъектBatchRequestRepeatCellФорматЯчеек()

Функция ПолучитьОбъектBatchRequestRepeatCellОграничениеРедактирования(ФабрикаGSs, ЛистИД, ПерваяКолонка = Неопределено, ПоследняяКолонка = Неопределено, 
		ПерваяСтрока = Неопределено, ПоследняяСтрока = Неопределено, СписокВыбора) Экспорт

	ТипBatchRequestRepeatCell	 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "BatchRequestRepeatCell");
	ТипRepeatCell				 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "RepeatCell");
	ТипGridRange				 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "GridRange");
	ТипCellData					 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "CellData");
	ТипDataValidationRule		 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "DataValidationRule");
	ТипBooleanCondition			 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "BooleanCondition");
	ТипConditionValue			 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "ConditionValue");
	
	ОбъектBatchRequestRepeatCell = ФабрикаGSs.Создать(ТипBatchRequestRepeatCell);
	
	ОбъектRepeatCell = ФабрикаGSs.Создать(ТипRepeatCell);
	
	ОбъектGridRange = ФабрикаGSs.Создать(ТипGridRange);
	
	ОбъектGridRange.sheetId			 = ЛистИД;
	
	Если НЕ ПерваяКолонка = Неопределено Тогда
		ОбъектGridRange.startColumnIndex = ПерваяКолонка - 1;
	КонецЕсли; 
	Если НЕ ПоследняяКолонка = Неопределено Тогда
		ОбъектGridRange.endColumnIndex = ПоследняяКолонка;
	КонецЕсли;
	Если НЕ ПерваяСтрока = Неопределено Тогда
		ОбъектGridRange.startRowIndex = ПерваяСтрока - 1;
	КонецЕсли;
	Если НЕ ПоследняяСтрока = Неопределено Тогда
		ОбъектGridRange.endRowIndex = ПоследняяСтрока;
	КонецЕсли;
	
	ОбъектRepeatCell.range = ОбъектGridRange;

	ОбъектCellData = ФабрикаGSs.Создать(ТипCellData);
	
	ОбъектDataValidationRule = ФабрикаGSs.Создать(ТипDataValidationRule);
	
	ОбъектBooleanCondition = ФабрикаGSs.Создать(ТипBooleanCondition);
	
	ОбъектBooleanCondition.type = "ONE_OF_LIST";
	
	ДоступныеЗначения = "";
	
	Для каждого ЗначениеВыбора Из СписокВыбора Цикл
		
		ОбъектConditionValue = ФабрикаGSs.Создать(ТипConditionValue);
		ОбъектConditionValue.userEnteredValue = ЗначениеВыбора;
		
		ОбъектBooleanCondition.values.Добавить(ОбъектConditionValue);
		
		ДоступныеЗначения = ?(НЕ ЗначениеЗаполнено(ДоступныеЗначения), Строка(ЗначениеВыбора), ДоступныеЗначения + ", " +  Строка(ЗначениеВыбора));
		
	КонецЦикла; 
	
	ОбъектDataValidationRule.condition = ОбъектBooleanCondition;
	ОбъектDataValidationRule.strict = "TRUE";
	
	ДоступныеЗначенияШаблон = НСтр("ru = 'Доступные значения для выбора: %1'");
	ДоступныеЗначения = СтрШаблон(ДоступныеЗначенияШаблон, ДоступныеЗначения);
	ОбъектDataValidationRule.inputMessage = ДоступныеЗначения;
	
	ОбъектCellData.dataValidation = ОбъектDataValidationRule;
	
	ОбъектRepeatCell.fields = "dataValidation";
	
	ОбъектRepeatCell.cell = ОбъектCellData;
	
	ОбъектBatchRequestRepeatCell.repeatCell = ОбъектRepeatCell;
	
	Возврат ОбъектBatchRequestRepeatCell;
	
КонецФункции // ПолучитьОбъектBatchRequestRepeatCell()

Функция ПолучитьОбъектBatchRequestRepeatCellОграничениеРедактированияТипЯчейкиЧисло(ФабрикаGSs, ЛистИД, ПерваяКолонка = Неопределено, ПоследняяКолонка = Неопределено, 
		ПерваяСтрока = Неопределено, ПоследняяСтрока = Неопределено) Экспорт

	ТипBatchRequestRepeatCell	 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "BatchRequestRepeatCell");
	ТипRepeatCell				 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "RepeatCell");
	ТипGridRange				 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "GridRange");
	ТипCellData					 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "CellData");
	ТипDataValidationRule		 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "DataValidationRule");
	ТипBooleanCondition			 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "BooleanCondition");
	ТипConditionValue			 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "ConditionValue");
	
	ОбъектBatchRequestRepeatCell = ФабрикаGSs.Создать(ТипBatchRequestRepeatCell);
	
	ОбъектRepeatCell = ФабрикаGSs.Создать(ТипRepeatCell);
	
	ОбъектGridRange = ФабрикаGSs.Создать(ТипGridRange);
	
	ОбъектGridRange.sheetId			 = ЛистИД;
	
	Если НЕ ПерваяКолонка = Неопределено Тогда
		ОбъектGridRange.startColumnIndex = ПерваяКолонка - 1;
	КонецЕсли; 
	Если НЕ ПоследняяКолонка = Неопределено Тогда
		ОбъектGridRange.endColumnIndex = ПоследняяКолонка;
	КонецЕсли;
	Если НЕ ПерваяСтрока = Неопределено Тогда
		ОбъектGridRange.startRowIndex = ПерваяСтрока - 1;
	КонецЕсли;
	Если НЕ ПоследняяСтрока = Неопределено Тогда
		ОбъектGridRange.endRowIndex = ПоследняяСтрока;
	КонецЕсли;
	
	ОбъектRepeatCell.range = ОбъектGridRange;

	ОбъектCellData = ФабрикаGSs.Создать(ТипCellData);
	
	ОбъектDataValidationRule = ФабрикаGSs.Создать(ТипDataValidationRule);
	
	ОбъектBooleanCondition = ФабрикаGSs.Создать(ТипBooleanCondition);
	
	ОбъектBooleanCondition.type = "NUMBER_GREATER_THAN_EQ";
	
	ОбъектConditionValue = ФабрикаGSs.Создать(ТипConditionValue);
	ОбъектConditionValue.userEnteredValue = "0";
		
	ОбъектBooleanCondition.values.Добавить(ОбъектConditionValue);
	
	ОбъектDataValidationRule.condition = ОбъектBooleanCondition;
	ОбъектDataValidationRule.strict = "TRUE";
	
	ОбъектDataValidationRule.inputMessage = "Доступен ввод только числового значения.";
	
	ОбъектCellData.dataValidation = ОбъектDataValidationRule;
	
	ОбъектRepeatCell.fields = "dataValidation";
	
	ОбъектRepeatCell.cell = ОбъектCellData;
	
	ОбъектBatchRequestRepeatCell.repeatCell = ОбъектRepeatCell;
	
	Возврат ОбъектBatchRequestRepeatCell;
	
КонецФункции // ПолучитьОбъектBatchRequestRepeatCell()

Функция ПолучитьОбъектAddProtectedRangeRequestУстановкаЗащитыНаЛист(ФабрикаGSs, ЛистИД, РедактируемыеОбласти, СлужебнаяУчетнаяЗапись) Экспорт

	ТипAddProtectedRangeRequests = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "AddProtectedRangeRequests");
	ТипAddProtectedRangeRequest	 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "AddProtectedRangeRequest");
	ТипProtectedRange			 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "ProtectedRange");
	ТипGridRange				 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "GridRange");
	ТипEditors					 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "Editors");
	
	ОбъектAddProtectedRangeRequests = ФабрикаGSs.Создать(ТипAddProtectedRangeRequests);
	
	ОбъектAddProtectedRangeRequest = ФабрикаGSs.Создать(ТипAddProtectedRangeRequest);
	
	ОбъектProtectedRange = ФабрикаGSs.Создать(ТипProtectedRange);
	
	ОбъектGridRange = ФабрикаGSs.Создать(ТипGridRange);
	
	ОбъектGridRange.sheetId			 = ЛистИД;
	
	ОбъектProtectedRange.range = ОбъектGridRange;
	
	Для каждого РедактируемаяОбласть Из РедактируемыеОбласти Цикл
		ОбъектProtectedRange.unprotectedRanges.Добавить(РедактируемаяОбласть);
	КонецЦикла;
	
	ОбъектEditors = ФабрикаGSs.Создать(ТипEditors);
	
	ОбъектEditors.users.Добавить(СлужебнаяУчетнаяЗапись);
	
	ОбъектProtectedRange.editors = ОбъектEditors;
	
	ОбъектAddProtectedRangeRequest.protectedRange = ОбъектProtectedRange;
	
	ОбъектAddProtectedRangeRequests.addProtectedRange = ОбъектAddProtectedRangeRequest;
	
	Возврат ОбъектAddProtectedRangeRequests;
	
КонецФункции // ПолучитьОбъектBatchRequestRepeatCell()

Функция ПолучитьОбъектPermissions(ФабрикаGSs, ПараметрыДоступаGSs)

	ТипPermissions = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "Permissions");
	ОбъектPermissions = ФабрикаGSs.Создать(ТипPermissions);
	
	ОбъектPermissions.emailAddress	 = ПараметрыДоступаGSs.emailAddress;
	ОбъектPermissions.role			 = ПараметрыДоступаGSs.role;
	ОбъектPermissions.type			 = ПараметрыДоступаGSs.type;
	ОбъектPermissions.sendNotificationEmail	 = ПараметрыДоступаGSs.sendNotificationEmail;

	Возврат ОбъектPermissions;
	
КонецФункции // ПолучитьОбъектPermissions()

////////////////////////////////////////////////////////////////////////////////
// Функции корректировки объектов JSON для корректной обработки платформой
//

Функция ПолучитьЗначениеОбъектаJSON(ОбъектJSON, УдалитьТабуляциюВНачалеСтроки = Истина) Экспорт 

	НовыйОбъектJSON = "";
	
	ОбъектJSONКолСтрок = СтрЧислоСтрок(ОбъектJSON);
	
	Для Сч = 1 По ОбъектJSONКолСтрок Цикл
		
		//..отсекаем первую и последнюю строку
		Если Сч = 1 ИЛИ Сч = ОбъектJSONКолСтрок Тогда
			Продолжить;
		КонецЕсли; 
		
		стрНовыйОбъектJSON = СтрПолучитьСтроку(ОбъектJSON,Сч);
		
		//..удаляем свойство value и незначащие символы
		Если СтрНайти(стрНовыйОбъектJSON, "#value") > 0 Тогда
		
			стрНовыйОбъектJSON = СтрЗаменить(стрНовыйОбъектJSON,"#value","");
			стрНовыйОбъектJSON = СтрЗаменить(стрНовыйОбъектJSON,"""","");
			стрНовыйОбъектJSON = СтрЗаменить(стрНовыйОбъектJSON," ","");
			стрНовыйОбъектJSON = СтрЗаменить(стрНовыйОбъектJSON,":","");
			
		КонецЕсли; 
		
		Если УдалитьТабуляциюВНачалеСтроки И Лев(стрНовыйОбъектJSON,1) = Символы.Таб Тогда 
			стрНовыйОбъектJSON = Прав(стрНовыйОбъектJSON,СтрДлина(стрНовыйОбъектJSON)-1);
		КонецЕсли;
		
		НовыйОбъектJSON = НовыйОбъектJSON + стрНовыйОбъектJSON + Символы.ПС;
		
	КонецЦикла;
	
	Возврат НовыйОбъектJSON;

КонецФункции // ПолучитьЗначениеОбъекаJSON()

Функция СформироватьЗначениеОбъектаJSON(ОбъектJSON) Экспорт 

	ЗначениеОбъектаJSON = "";
	
	ОбъектJSONКолСтрок = СтрЧислоСтрок(ОбъектJSON);
		
	ЗначениеОбъектаJSONШаблон = "{
			|""#value"": %1}";
	
	ЗначениеОбъектаJSON = СтрШаблон(ЗначениеОбъектаJSONШаблон, ОбъектJSON); 
	
	Возврат ЗначениеОбъектаJSON;

КонецФункции // ПолучитьЗначениеОбъекаJSON()

//функция выполняет последовательное построчное чтение ОбъектJSON
//свойство должно быть уникальным в объекте!
Функция ПрочитатьСвойствоВСтрокеОбъектаJSON(ОбъектJSON, ИмяСвойства) Экспорт 

	ЗначениеСвойства = Неопределено;

	ЧтениеJSON = Новый ЧтениеJSON;
	ЧтениеJSON.УстановитьСтроку(ОбъектJSON);
	
	Пока ЧтениеJSON.Прочитать() Цикл
		
		Если ЧтениеJSON.ТипТекущегоЗначения = ТипЗначенияJSON.ИмяСвойства 
			И ЧтениеJSON.ТекущееЗначение = ИмяСвойства Тогда
			
			ЧтениеJSON.Прочитать();

			ЗначениеСвойства = ЧтениеJSON.ТекущееЗначение;
			
			Прервать;
			
		КонецЕсли; 
		
	КонецЦикла;
 	
	Возврат ЗначениеСвойства;
	
КонецФункции // ПрочитатьСвойствоВСтрокеОбъектаJSON()

Функция ПолучитьОбъектОбъединениеЯчеек(ФабрикаGSs, ЛистИД, ПерваяКолонка, ПоследняяКолонка, ПерваяСтрока, ПоследняяСтрока, ТипОбъединения = "MERGE_ALL") Экспорт 
	
	ТипBatchRequestMergeCell = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "BatchRequestMergeCell");
	ТипMergeCells			 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "MergeCells");
	ТипGridRange			 = ФабрикаGSs.Тип("http://www.googlespreadsheet.org", "GridRange");
	
	ОбъектBatchRequestMergeCell = ФабрикаGSs.Создать(ТипBatchRequestMergeCell);
	
	ОбъектMergeCells = ФабрикаGSs.Создать(ТипMergeCells);
	
	ОбъектGridRange = ФабрикаGSs.Создать(ТипGridRange);
	
	ОбъектGridRange.sheetId			 = ЛистИД;
	ОбъектGridRange.startColumnIndex = ПерваяКолонка - 1;
	ОбъектGridRange.endColumnIndex	 = ПоследняяКолонка;
	ОбъектGridRange.startRowIndex	 = ПерваяСтрока - 1;
	ОбъектGridRange.endRowIndex		 = ПоследняяСтрока;
	
	ОбъектMergeCells.range = ОбъектGridRange;
	
	ОбъектMergeCells.mergeType = ТипОбъединения;
	
	ОбъектBatchRequestMergeCell.mergeCells = ОбъектMergeCells;
	
	Возврат ОбъектBatchRequestMergeCell;
	
КонецФункции

Процедура ЗаполнитьПоУмолчаниюParentsFile(ОбъектGSs, ФабрикаGSs, ДопПараметры) 

	ТипDriveProperties = ФабрикаGSs.Тип("http://www.GoogleAPI_GoogleDrive.org", "ParentsFile");
	ОбъектDriveProperties = ФабрикаGSs.Создать(ТипDriveProperties);
	
	GoogleAPI_СериализацияОбъектовGoogleJSON.ЗаполнитьОбъектПоУмолчаниюGSs(ОбъектDriveProperties, ФабрикаGSs, ДопПараметры);
	
	ОбъектDriveProperties.id 		= ДопПараметры.parents_id;
	ОбъектDriveProperties.isRoot 	= false;
	ОбъектDriveProperties.kind 		= "";
	ОбъектDriveProperties.parentLink 	= "";
	ОбъектDriveProperties.selfLink 	= "";

	ОбъектGSs.parents.Добавить(ОбъектDriveProperties);
	
КонецПроцедуры // ЗаполнитьПоУмолчаниюSheet()

#КонецОбласти




unit adddata;

interface

uses Classes, SysUtils, DateUtils, VCL.FlexCel.Core, FlexCel.XlsAdapter;

procedure CreateAndSaveDataFile;
procedure CreateExcelDataFile(const xls: TExcelFile);


implementation

uses Main;

procedure CreateAndSaveDataFile;
var
  xls: TXlsFile;
begin
  xls := TXlsFile.Create(true);
  try
    CreateExcelDataFile(xls);
    //Save the file as XLS
    xls.Save(dataFileName);
  finally
    xls.Free;
  end
end;

procedure CreateExcelDataFile(const xls: TExcelFile);
var
  StyleFmt: TFlxFormat;
  MajorLatin: TThemeTextFont;
  MajorEastAsian: TThemeTextFont;
  MajorComplexScript: TThemeTextFont;
  MajorFont: TThemeFont;
  MinorLatin: TThemeTextFont;
  MinorEastAsian: TThemeTextFont;
  MinorComplexScript: TThemeTextFont;
  MinorFont: TThemeFont;
  fmt: TFlxFormat;
  Runs: TArray<TRTFRun>;
  fnt: TFlxFont;
  Link: THyperLink;

begin
  xls.NewFile(2, TExcelFileFormat.v2019);  //Create a new Excel file with 2 sheets.

  //Set the names of the sheets
  xls.ActiveSheet := 1;
  xls.SheetName := 'Data';
  xls.ActiveSheet := 2;
  xls.SheetName := 'Project documents';

  xls.ActiveSheet := 1;  //Set the sheet we are working in.

  //Global Workbook Options
  //Note that in xlsx files this option is ignored by FlexCel unless you also set OptionsForceUseCheckCompatibility to true. This is because Excel disables Autosave in files which have this option.
  xls.OptionsCheckCompatibility := false;

  //Sheet Options
  xls.SheetName := 'Data';

  //Styles.
  StyleFmt := xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Hyperlink, 0));
  StyleFmt.VAlignment := TVFlxAlignment.bottom;
  StyleFmt.Locked := true;
  xls.SetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Hyperlink, 0), StyleFmt);

  //Printer Settings
  xls.PrintXResolution := 600;
  xls.PrintYResolution := 600;
  xls.PrintOptions := [TPrintOptions.Orientation];
  xls.PrintPaperSize := TPaperSize.A4;

  //Theme - You might use GetTheme/SetTheme methods here instead.
  xls.SetColorTheme(TPrimaryThemeColor.Accent1, TDrawingColor.FromRgb($5B, $9B, $D5));
  xls.SetColorTheme(TPrimaryThemeColor.Accent5, TDrawingColor.FromRgb($44, $72, $C4));

  //Major font
  MajorLatin := TThemeTextFont.Create('Calibri Light', '020F0302020204030204', TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
  MajorEastAsian := TThemeTextFont.Create('', '', TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
  MajorComplexScript := TThemeTextFont.Create('', '', TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
  MajorFont := TThemeFont.Create(MajorLatin, MajorEastAsian, MajorComplexScript);
  xls.SetThemeFont(TFontScheme.Major, MajorFont);

  //Minor font
  MinorLatin := TThemeTextFont.Create('Calibri', '020F0502020204030204', TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
  MinorEastAsian := TThemeTextFont.Create('', '', TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
  MinorComplexScript := TThemeTextFont.Create('', '', TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
  MinorFont := TThemeFont.Create(MinorLatin, MinorEastAsian, MinorComplexScript);
  xls.SetThemeFont(TFontScheme.Minor, MinorFont);

  //Set up rows and columns
  xls.SetColWidth(1, 11, 5778);  //(21.82 + 0.75) * 256

  xls.SetColWidth(12, 12, 4096);  //(15.25 + 0.75) * 256
  xls.DefaultRowHeight := 300;

  xls.SetRowHeight(1, 960);  //48.00 * 20
  xls.SetRowHeight(3, 315);  //15.75 * 20
  xls.SetRowHeight(4, 315);  //15.75 * 20
  xls.SetRowHeight(5, 315);  //15.75 * 20
  xls.SetRowHeight(6, 315);  //15.75 * 20
  xls.SetRowHeight(7, 315);  //15.75 * 20
  xls.SetRowHeight(8, 315);  //15.75 * 20
  xls.SetRowHeight(11, 315);  //15.75 * 20
  xls.SetRowHeight(12, 315);  //15.75 * 20
  xls.SetRowHeight(13, 315);  //15.75 * 20
  xls.SetRowHeight(14, 315);  //15.75 * 20
  xls.SetRowHeight(15, 315);  //15.75 * 20
  xls.SetRowHeight(16, 315);  //15.75 * 20
  xls.SetRowHeight(17, 315);  //15.75 * 20

  //Merged Cells
  xls.MergeCells(1, 1, 1, 9);
  xls.MergeCells(5, 1, 5, 9);

  //Set the cell values
  fmt := xls.GetCellVisibleFormatDef(1, 1);
  fmt.Font.Name := 'Calibri Light';
  fmt.Font.Size20 := 720;
  fmt.Font.Color := TExcelColor.FromTheme(TThemeColor.Light1);
  fmt.Font.Scheme := TFontScheme.Major;
  fmt.FillPattern.Pattern := TFlxPatternStyle.Solid;
  fmt.FillPattern.FgColor := TExcelColor.FromTheme(TThemeColor.Accent1, -0.499984740745262);
  fmt.FillPattern.BgColor := TExcelColor.Automatic;
  fmt.HAlignment := THFlxAlignment.left;
  fmt.VAlignment := TVFlxAlignment.center;
  xls.SetCellFormat(1, 1, xls.AddFormat(fmt));
  xls.SetCellValue(1, 1, 'Source data for "Select data element" drop-down menu');

  fmt := xls.GetCellVisibleFormatDef(1, 2);
  fmt.Font.Name := 'Calibri Light';
  fmt.Font.Size20 := 720;
  fmt.Font.Color := TExcelColor.FromTheme(TThemeColor.Light1);
  fmt.Font.Scheme := TFontScheme.Major;
  fmt.FillPattern.Pattern := TFlxPatternStyle.Solid;
  fmt.FillPattern.FgColor := TExcelColor.FromTheme(TThemeColor.Accent1, -0.499984740745262);
  fmt.FillPattern.BgColor := TExcelColor.Automatic;
  fmt.HAlignment := THFlxAlignment.left;
  fmt.VAlignment := TVFlxAlignment.center;
  xls.SetCellFormat(1, 2, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(1, 3);
  fmt.Font.Name := 'Calibri Light';
  fmt.Font.Size20 := 720;
  fmt.Font.Color := TExcelColor.FromTheme(TThemeColor.Light1);
  fmt.Font.Scheme := TFontScheme.Major;
  fmt.FillPattern.Pattern := TFlxPatternStyle.Solid;
  fmt.FillPattern.FgColor := TExcelColor.FromTheme(TThemeColor.Accent1, -0.499984740745262);
  fmt.FillPattern.BgColor := TExcelColor.Automatic;
  fmt.HAlignment := THFlxAlignment.left;
  fmt.VAlignment := TVFlxAlignment.center;
  xls.SetCellFormat(1, 3, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(1, 4);
  fmt.Font.Name := 'Calibri Light';
  fmt.Font.Size20 := 720;
  fmt.Font.Color := TExcelColor.FromTheme(TThemeColor.Light1);
  fmt.Font.Scheme := TFontScheme.Major;
  fmt.FillPattern.Pattern := TFlxPatternStyle.Solid;
  fmt.FillPattern.FgColor := TExcelColor.FromTheme(TThemeColor.Accent1, -0.499984740745262);
  fmt.FillPattern.BgColor := TExcelColor.Automatic;
  fmt.HAlignment := THFlxAlignment.left;
  fmt.VAlignment := TVFlxAlignment.center;
  xls.SetCellFormat(1, 4, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(1, 5);
  fmt.Font.Name := 'Calibri Light';
  fmt.Font.Size20 := 720;
  fmt.Font.Color := TExcelColor.FromTheme(TThemeColor.Light1);
  fmt.Font.Scheme := TFontScheme.Major;
  fmt.FillPattern.Pattern := TFlxPatternStyle.Solid;
  fmt.FillPattern.FgColor := TExcelColor.FromTheme(TThemeColor.Accent1, -0.499984740745262);
  fmt.FillPattern.BgColor := TExcelColor.Automatic;
  fmt.HAlignment := THFlxAlignment.left;
  fmt.VAlignment := TVFlxAlignment.center;
  xls.SetCellFormat(1, 5, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(1, 6);
  fmt.Font.Name := 'Calibri Light';
  fmt.Font.Size20 := 720;
  fmt.Font.Color := TExcelColor.FromTheme(TThemeColor.Light1);
  fmt.Font.Scheme := TFontScheme.Major;
  fmt.FillPattern.Pattern := TFlxPatternStyle.Solid;
  fmt.FillPattern.FgColor := TExcelColor.FromTheme(TThemeColor.Accent1, -0.499984740745262);
  fmt.FillPattern.BgColor := TExcelColor.Automatic;
  fmt.HAlignment := THFlxAlignment.left;
  fmt.VAlignment := TVFlxAlignment.center;
  xls.SetCellFormat(1, 6, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(1, 7);
  fmt.Font.Name := 'Calibri Light';
  fmt.Font.Size20 := 720;
  fmt.Font.Color := TExcelColor.FromTheme(TThemeColor.Light1);
  fmt.Font.Scheme := TFontScheme.Major;
  fmt.FillPattern.Pattern := TFlxPatternStyle.Solid;
  fmt.FillPattern.FgColor := TExcelColor.FromTheme(TThemeColor.Accent1, -0.499984740745262);
  fmt.FillPattern.BgColor := TExcelColor.Automatic;
  fmt.HAlignment := THFlxAlignment.left;
  fmt.VAlignment := TVFlxAlignment.center;
  xls.SetCellFormat(1, 7, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(1, 8);
  fmt.Font.Name := 'Calibri Light';
  fmt.Font.Size20 := 720;
  fmt.Font.Color := TExcelColor.FromTheme(TThemeColor.Light1);
  fmt.Font.Scheme := TFontScheme.Major;
  fmt.FillPattern.Pattern := TFlxPatternStyle.Solid;
  fmt.FillPattern.FgColor := TExcelColor.FromTheme(TThemeColor.Accent1, -0.499984740745262);
  fmt.FillPattern.BgColor := TExcelColor.Automatic;
  fmt.HAlignment := THFlxAlignment.left;
  fmt.VAlignment := TVFlxAlignment.center;
  xls.SetCellFormat(1, 8, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(1, 9);
  fmt.Font.Name := 'Calibri Light';
  fmt.Font.Size20 := 720;
  fmt.Font.Color := TExcelColor.FromTheme(TThemeColor.Light1);
  fmt.Font.Scheme := TFontScheme.Major;
  fmt.FillPattern.Pattern := TFlxPatternStyle.Solid;
  fmt.FillPattern.FgColor := TExcelColor.FromTheme(TThemeColor.Accent1, -0.499984740745262);
  fmt.FillPattern.BgColor := TExcelColor.Automatic;
  fmt.HAlignment := THFlxAlignment.left;
  fmt.VAlignment := TVFlxAlignment.center;
  xls.SetCellFormat(1, 9, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(3, 1);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(3, 1, xls.AddFormat(fmt));

  SetLength(Runs, 2);
  Runs[0].FirstChar := 10;
  fnt := xls.GetDefaultFont;
  fnt.Size20 := 240;
  fnt.Color := TExcelColor.Automatic;
  Runs[0].Font := fnt;
  Runs[1].FirstChar := 15;
  fnt := xls.GetDefaultFont;
  fnt.Size20 := 240;
  Runs[1].Font := fnt;
  xls.SetCellValue(3, 1, TRichString.Create('Variables names (row 10) will be read for each column, starting at A10, until it'
  + ' hits an empty cell.', Runs));
  //We could also have used: xls.SetCellFromHtml(3, 1, 'Variables&nbsp;<font color = ''black''>names</font>&nbsp;(row 10) will be read for'
  //
  //  + ' each column, starting at A10, until it hits an empty cell.');


  fmt := xls.GetCellVisibleFormatDef(3, 2);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(3, 2, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(3, 3);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(3, 3, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(3, 4);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(3, 4, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(3, 5);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(3, 5, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(3, 6);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(3, 6, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(3, 7);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(3, 7, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(3, 8);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(3, 8, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(3, 9);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(3, 9, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(4, 1);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(4, 1, xls.AddFormat(fmt));

  SetLength(Runs, 2);
  Runs[0].FirstChar := 9;
  fnt := xls.GetDefaultFont;
  fnt.Size20 := 240;
  fnt.Color := TExcelColor.Automatic;
  Runs[0].Font := fnt;
  Runs[1].FirstChar := 13;
  fnt := xls.GetDefaultFont;
  fnt.Size20 := 240;
  Runs[1].Font := fnt;
  xls.SetCellValue(4, 1, TRichString.Create('Variable data in a cell (row 11 and down), can be empty, except for the "Dropdown'
  + ' Title".', Runs));
  //We could also have used: xls.SetCellFromHtml(4, 1, 'Variable&nbsp;<font color = ''black''>data</font>&nbsp;in a cell (row 11 and down),'
  //
  //  + ' can be empty, except for the &quot;Dropdown Title&quot;.');


  fmt := xls.GetCellVisibleFormatDef(4, 2);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(4, 2, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(4, 3);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(4, 3, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(4, 4);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(4, 4, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(4, 5);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(4, 5, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(4, 6);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(4, 6, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(4, 7);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(4, 7, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(4, 8);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(4, 8, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(4, 9);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(4, 9, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(5, 1);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(5, 1, xls.AddFormat(fmt));
  xls.SetCellValue(5, 1, 'Application will read each row for data, starting at row 11, until it hits an empty'
  + ' cell.');

  fmt := xls.GetCellVisibleFormatDef(5, 2);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(5, 2, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(5, 3);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(5, 3, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(5, 4);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(5, 4, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(5, 5);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(5, 5, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(5, 6);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(5, 6, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(5, 7);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(5, 7, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(5, 8);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(5, 8, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(5, 9);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(5, 9, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(6, 1);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(6, 1, xls.AddFormat(fmt));
  xls.SetCellValue(6, 1, 'You can create as many variables and rows as you want, just expand the table. Make'
  + ' sure to use unique variable names.');

  fmt := xls.GetCellVisibleFormatDef(6, 2);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(6, 2, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(6, 3);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(6, 3, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(6, 4);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(6, 4, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(6, 5);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(6, 5, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(6, 6);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(6, 6, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(6, 7);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(6, 7, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(6, 8);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(6, 8, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(6, 9);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(6, 9, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(7, 1);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(7, 1, xls.AddFormat(fmt));
  xls.SetCellValue(7, 1, 'On selection from drop-down menu in app, the application will assign data for that'
  + ' selection into each variable');

  fmt := xls.GetCellVisibleFormatDef(7, 2);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(7, 2, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(7, 3);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(7, 3, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(7, 4);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(7, 4, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(7, 5);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(7, 5, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(7, 6);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(7, 6, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(7, 7);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(7, 7, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(7, 8);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(7, 8, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(7, 9);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(7, 9, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(8, 1);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(8, 1, xls.AddFormat(fmt));
  xls.SetCellValue(8, 1, 'During script execution, the script and any content will be string replaced with'
  + ' all variables and their values');

  fmt := xls.GetCellVisibleFormatDef(8, 2);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(8, 2, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(8, 3);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(8, 3, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(8, 4);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(8, 4, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(8, 5);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(8, 5, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(8, 6);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(8, 6, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(8, 7);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(8, 7, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(8, 8);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(8, 8, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(8, 9);
  fmt.Font.Size20 := 240;
  xls.SetCellFormat(8, 9, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(10, 1);
  fmt.Font.Style := [TFlxFontStyles.Bold];
  xls.SetCellFormat(10, 1, xls.AddFormat(fmt));
  xls.SetCellValue(10, 1, 'Dropdown Title');
  xls.SetCellValue(10, 2, '#site-id#');

  fmt := xls.GetCellVisibleFormatDef(10, 3);
  fmt.Font.Style := [TFlxFontStyles.Bold];
  xls.SetCellFormat(10, 3, xls.AddFormat(fmt));
  xls.SetCellValue(10, 3, '#ipv4-block#');

  fmt := xls.GetCellVisibleFormatDef(10, 4);
  fmt.Font.Style := [TFlxFontStyles.Bold];
  xls.SetCellFormat(10, 4, xls.AddFormat(fmt));
  xls.SetCellValue(10, 4, '#ipv4-block-mask#');

  fmt := xls.GetCellVisibleFormatDef(10, 5);
  fmt.Font.Style := [TFlxFontStyles.Bold];
  xls.SetCellFormat(10, 5, xls.AddFormat(fmt));
  xls.SetCellValue(10, 5, '#ipv4-mgmt-net#');

  fmt := xls.GetCellVisibleFormatDef(10, 6);
  fmt.Font.Style := [TFlxFontStyles.Bold];
  xls.SetCellFormat(10, 6, xls.AddFormat(fmt));
  xls.SetCellValue(10, 6, '#ipv4-mgmt-mask#');

  fmt := xls.GetCellVisibleFormatDef(10, 7);
  fmt.Font.Style := [TFlxFontStyles.Bold];
  xls.SetCellFormat(10, 7, xls.AddFormat(fmt));
  xls.SetCellValue(10, 7, '#ipv4-mgmt-gw#');

  fmt := xls.GetCellVisibleFormatDef(10, 8);
  fmt.Font.Style := [TFlxFontStyles.Bold];
  xls.SetCellFormat(10, 8, xls.AddFormat(fmt));
  xls.SetCellValue(10, 8, '#ipv6-block#');

  fmt := xls.GetCellVisibleFormatDef(10, 9);
  fmt.Font.Style := [TFlxFontStyles.Bold];
  xls.SetCellFormat(10, 9, xls.AddFormat(fmt));
  xls.SetCellValue(10, 9, '#ipv6-block-mask#');

  fmt := xls.GetCellVisibleFormatDef(10, 10);
  fmt.Font.Style := [TFlxFontStyles.Bold];
  xls.SetCellFormat(10, 10, xls.AddFormat(fmt));
  xls.SetCellValue(10, 10, '#ipv6-mgmt-net#');

  fmt := xls.GetCellVisibleFormatDef(10, 11);
  fmt.Font.Style := [TFlxFontStyles.Bold];
  xls.SetCellFormat(10, 11, xls.AddFormat(fmt));
  xls.SetCellValue(10, 11, '#ipv6-mgmt-mask#');

  fmt := xls.GetCellVisibleFormatDef(10, 12);
  fmt.Font.Style := [TFlxFontStyles.Bold];
  xls.SetCellFormat(10, 12, xls.AddFormat(fmt));
  xls.SetCellValue(10, 12, '#ipv6-mgmt-gw#');

  fmt := xls.GetCellVisibleFormatDef(11, 1);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(11, 1, xls.AddFormat(fmt));
  xls.SetCellValue(11, 1, 'Bergen-Office');
  xls.SetCellValue(11, 2, '1f100');

  fmt := xls.GetCellVisibleFormatDef(11, 3);
  fmt.Format := 'mmm-yy';
  xls.SetCellFormat(11, 3, xls.AddFormat(fmt));
  xls.SetCellValue(11, 3, '10.80.');
  xls.SetCellValue(11, 4, '/16');
  xls.SetCellValue(11, 5, '10.80.10.');
  xls.SetCellValue(11, 6, '/24');
  xls.SetCellValue(11, 7, '10.80.10.1');
  xls.SetCellValue(11, 8, 'fd00:1:f100::');
  xls.SetCellValue(11, 9, '/48');
  xls.SetCellValue(11, 10, 'fd00:1:f100:10::');
  xls.SetCellValue(11, 11, '/64');
  xls.SetCellValue(11, 12, 'fd00:1:f100:10::1');

  fmt := xls.GetCellVisibleFormatDef(12, 1);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(12, 1, xls.AddFormat(fmt));
  xls.SetCellValue(12, 1, 'London-Office');
  xls.SetCellValue(12, 2, '1f101');
  xls.SetCellValue(12, 3, '10.81.');
  xls.SetCellValue(12, 4, '/16');
  xls.SetCellValue(12, 5, '10.81.10');
  xls.SetCellValue(12, 6, '/24');
  xls.SetCellValue(12, 7, '10.81.10.1');
  xls.SetCellValue(12, 8, 'fd00:1:f101::');
  xls.SetCellValue(12, 9, '/48');
  xls.SetCellValue(12, 10, 'fd00:1:f101:10::');
  xls.SetCellValue(12, 11, '/64');
  xls.SetCellValue(12, 12, 'fd00:1:f101:10::1');

  fmt := xls.GetCellVisibleFormatDef(13, 1);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(13, 1, xls.AddFormat(fmt));
  xls.SetCellValue(13, 1, 'Amsterdam-Office');
  xls.SetCellValue(13, 2, '1f102');

  fmt := xls.GetCellVisibleFormatDef(13, 3);
  fmt.Format := 'mmm-yy';
  xls.SetCellFormat(13, 3, xls.AddFormat(fmt));
  xls.SetCellValue(13, 3, '10.82.');
  xls.SetCellValue(13, 4, '/16');
  xls.SetCellValue(13, 5, '10.82.10.');
  xls.SetCellValue(13, 6, '/24');
  xls.SetCellValue(13, 7, '10.82.10.1');
  xls.SetCellValue(13, 8, 'fd00:1:f102::');
  xls.SetCellValue(13, 9, '/48');
  xls.SetCellValue(13, 10, 'fd00:1:f102:10::');
  xls.SetCellValue(13, 11, '/64');
  xls.SetCellValue(13, 12, 'fd00:1:f102:10::1');

  fmt := xls.GetCellVisibleFormatDef(14, 1);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(14, 1, xls.AddFormat(fmt));
  xls.SetCellValue(14, 1, 'New-York-Office');
  xls.SetCellValue(14, 2, '1f103');
  xls.SetCellValue(14, 3, '10.83.');
  xls.SetCellValue(14, 4, '/16');
  xls.SetCellValue(14, 5, '10.83.10');
  xls.SetCellValue(14, 6, '/24');
  xls.SetCellValue(14, 7, '10.83.10.1');
  xls.SetCellValue(14, 8, 'fd00:1:f103::');
  xls.SetCellValue(14, 9, '/48');
  xls.SetCellValue(14, 10, 'fd00:1:f103:10::');
  xls.SetCellValue(14, 11, '/64');
  xls.SetCellValue(14, 12, 'fd00:1:f103:10::1');

  fmt := xls.GetCellVisibleFormatDef(15, 1);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(15, 1, xls.AddFormat(fmt));
  xls.SetCellValue(15, 1, 'Rio-Office');
  xls.SetCellValue(15, 2, '1f104');

  fmt := xls.GetCellVisibleFormatDef(15, 3);
  fmt.Format := 'mmm-yy';
  xls.SetCellFormat(15, 3, xls.AddFormat(fmt));
  xls.SetCellValue(15, 3, '10.84.');
  xls.SetCellValue(15, 4, '/16');
  xls.SetCellValue(15, 5, '10.84.10.');
  xls.SetCellValue(15, 6, '/24');
  xls.SetCellValue(15, 7, '10.84.10.1');
  xls.SetCellValue(15, 8, 'fd00:1:f104:::');
  xls.SetCellValue(15, 9, '/48');
  xls.SetCellValue(15, 10, 'fd00:1:f104:10::');
  xls.SetCellValue(15, 11, '/64');
  xls.SetCellValue(15, 12, 'fd00:1:f104:10::1');

  fmt := xls.GetCellVisibleFormatDef(16, 1);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(16, 1, xls.AddFormat(fmt));
  xls.SetCellValue(16, 1, 'Perth-Office');
  xls.SetCellValue(16, 2, '1f105');
  xls.SetCellValue(16, 3, '10.85.');
  xls.SetCellValue(16, 4, '/16');
  xls.SetCellValue(16, 5, '10.85.10');
  xls.SetCellValue(16, 6, '/24');
  xls.SetCellValue(16, 7, '10.85.10.1');
  xls.SetCellValue(16, 8, 'fd00:1:f105:::');
  xls.SetCellValue(16, 9, '/48');
  xls.SetCellValue(16, 10, 'fd00:1:f105:10::');
  xls.SetCellValue(16, 11, '/64');
  xls.SetCellValue(16, 12, 'fd00:1:f105:10::1');

  fmt := xls.GetCellVisibleFormatDef(17, 1);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(17, 1, xls.AddFormat(fmt));
  xls.SetCellValue(17, 1, 'Tokyo-Office');
  xls.SetCellValue(17, 2, '1f106');

  fmt := xls.GetCellVisibleFormatDef(17, 3);
  fmt.Format := 'mmm-yy';
  xls.SetCellFormat(17, 3, xls.AddFormat(fmt));
  xls.SetCellValue(17, 3, '10.86.');
  xls.SetCellValue(17, 4, '/16');
  xls.SetCellValue(17, 5, '10.86.10');
  xls.SetCellValue(17, 6, '/24');
  xls.SetCellValue(17, 7, '10.86.10.1');
  xls.SetCellValue(17, 8, 'fd00:1:f106:::');
  xls.SetCellValue(17, 9, '/48');
  xls.SetCellValue(17, 10, 'fd00:1:f106:10::');
  xls.SetCellValue(17, 11, '/64');
  xls.SetCellValue(17, 12, 'fd00:1:f106:10::1');

  //Cell selection and scroll position.
  xls.SelectCell(1, 1, false);


  xls.ActiveSheet := 2;  //Set the sheet we are working in.

  //Sheet Options
  xls.SheetName := 'Project documents';

  //Set up rows and columns
  xls.SetColWidth(1, 1, 11958);  //(45.96 + 0.75) * 256

  xls.SetColWidth(2, 2, 30171);  //(117.11 + 0.75) * 256

  xls.SetRowHeight(1, 960);  //48.00 * 20
  xls.SetRowHeight(2, 315);  //15.75 * 20
  xls.SetRowHeight(3, 315);  //15.75 * 20
  xls.SetRowHeight(4, 315);  //15.75 * 20
  xls.SetRowHeight(5, 315);  //15.75 * 20

  //Merged Cells
  xls.MergeCells(1, 1, 1, 2);

  //Set the cell values
  fmt := xls.GetCellVisibleFormatDef(1, 1);
  fmt.Font.Name := 'Calibri Light';
  fmt.Font.Size20 := 720;
  fmt.Font.Color := TExcelColor.FromTheme(TThemeColor.Light1);
  fmt.Font.Scheme := TFontScheme.Major;
  fmt.FillPattern.Pattern := TFlxPatternStyle.Solid;
  fmt.FillPattern.FgColor := TExcelColor.FromTheme(TThemeColor.Accent1, -0.499984740745262);
  fmt.FillPattern.BgColor := TExcelColor.Automatic;
  fmt.HAlignment := THFlxAlignment.left;
  fmt.VAlignment := TVFlxAlignment.center;
  xls.SetCellFormat(1, 1, xls.AddFormat(fmt));
  xls.SetCellValue(1, 1, 'Project Documents Menu');

  fmt := xls.GetCellVisibleFormatDef(1, 2);
  fmt.Font.Name := 'Calibri Light';
  fmt.Font.Size20 := 720;
  fmt.Font.Color := TExcelColor.FromTheme(TThemeColor.Light1);
  fmt.Font.Scheme := TFontScheme.Major;
  fmt.FillPattern.Pattern := TFlxPatternStyle.Solid;
  fmt.FillPattern.FgColor := TExcelColor.FromTheme(TThemeColor.Accent1, -0.499984740745262);
  fmt.FillPattern.BgColor := TExcelColor.Automatic;
  fmt.HAlignment := THFlxAlignment.left;
  fmt.VAlignment := TVFlxAlignment.center;
  xls.SetCellFormat(1, 2, xls.AddFormat(fmt));

  fmt := xls.GetCellVisibleFormatDef(2, 1);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(2, 1, xls.AddFormat(fmt));
  xls.SetCellValue(2, 1, 'Hyperlinks to web sites and documents related to the project.');

  fmt := xls.GetCellVisibleFormatDef(3, 1);
  fmt.Font.Size20 := 240;
  fmt.HAlignment := THFlxAlignment.left;
  xls.SetCellFormat(3, 1, xls.AddFormat(fmt));
  xls.SetCellValue(3, 1, 'Start from row 6 and enter as many as needed. They will all show up in the "Project'
  + ' documents" menu');

  fmt := xls.GetCellVisibleFormatDef(5, 1);
  fmt.Font.Style := [TFlxFontStyles.Bold];
  fmt.Borders.Left.Style := TFlxBorderStyle.Medium;
  fmt.Borders.Left.Color := TExcelColor.Automatic;
  fmt.Borders.Top.Style := TFlxBorderStyle.Medium;
  fmt.Borders.Top.Color := TExcelColor.Automatic;
  fmt.Borders.Bottom.Style := TFlxBorderStyle.Medium;
  fmt.Borders.Bottom.Color := TExcelColor.Automatic;
  fmt.FillPattern.Pattern := TFlxPatternStyle.Solid;
  fmt.FillPattern.FgColor := TExcelColor.FromTheme(TThemeColor.Accent4, 0.799951170384838);
  fmt.FillPattern.BgColor := TExcelColor.Automatic;
  xls.SetCellFormat(5, 1, xls.AddFormat(fmt));
  xls.SetCellValue(5, 1, 'Menu description');

  fmt := xls.GetCellVisibleFormatDef(5, 2);
  fmt.Font.Style := [TFlxFontStyles.Bold];
  fmt.Borders.Right.Style := TFlxBorderStyle.Medium;
  fmt.Borders.Right.Color := TExcelColor.Automatic;
  fmt.Borders.Top.Style := TFlxBorderStyle.Medium;
  fmt.Borders.Top.Color := TExcelColor.Automatic;
  fmt.Borders.Bottom.Style := TFlxBorderStyle.Medium;
  fmt.Borders.Bottom.Color := TExcelColor.Automatic;
  fmt.FillPattern.Pattern := TFlxPatternStyle.Solid;
  fmt.FillPattern.FgColor := TExcelColor.FromTheme(TThemeColor.Accent4, 0.799951170384838);
  fmt.FillPattern.BgColor := TExcelColor.Automatic;
  xls.SetCellFormat(5, 2, xls.AddFormat(fmt));
  xls.SetCellValue(5, 2, 'Hyperlink, application or document');
  xls.SetCellValue(6, 1, 'Project documentation');

  fmt := xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Hyperlink, 0), true);
  xls.SetCellFormat(6, 2, xls.AddFormat(fmt));
  xls.SetCellValue(6, 2, 'https://www.company.com/project-documentation.html');
  xls.SetCellValue(7, 1, 'All your searching needs');

  fmt := xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Hyperlink, 0), true);
  xls.SetCellFormat(7, 2, xls.AddFormat(fmt));
  xls.SetCellValue(7, 2, 'https://www.google.com');
  xls.SetCellValue(8, 1, 'Someone to talk to');
  xls.SetCellValue(8, 2, 'https://chat.openai.com/');

  //Hyperlinks
  Link := THyperLink.Create(THyperLinkType.URL, 'https://www.company.com/project-documentation.html', '', '', '');
  xls.AddHyperLink(TXlsCellRange.Create(6, 2, 6, 2), Link);
  Link := THyperLink.Create(THyperLinkType.URL, 'https://www.google.com/', '', '', '');
  xls.AddHyperLink(TXlsCellRange.Create(7, 2, 7, 2), Link);

  //Cell selection and scroll position.
  xls.SelectCell(1, 1, false);

  //Standard Document Properties - Most are only for xlsx files. In xls files FlexCel will only change the Creation Date and Modified Date.
  xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, 'jorlan');

  xls.ActiveSheet := 1;

end;

end.

unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ADODB, DB, ComCtrls, IniFiles;

type
  TForm1 = class(TForm)
    btnProcess: TButton;
    btnCancel: TButton;
    MemoLog: TMemo;
    ProgressBar1: TProgressBar;
    OpenDialog1: TOpenDialog;
    procedure btnProcessClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    FConnection: TADOConnection;
    FQuery: TADOQuery;
    FCSVFileName: string;
    FTableName: string;
    FCancelRequested: Boolean;
    FConfigFileName: string;
    FConnectionString: string;

    //UI
    function GetCSVFile: Boolean;
    procedure UpdateProgress(Current, Total: Int64);

    //Logging
    procedure LogMessage(const Msg: string);

    //Config
    procedure LoadConfig;
    procedure CreateDefaultConfig;

    //DB Connection
    procedure InitConnection;
    procedure CleanupConnection;

    // DB Indexes
    procedure DropIndexIfExists(const TableName, IndexName: string);
    procedure CreateIndex(const TableName, IndexName, Columns: string);
    
    // DB Table Operations
    function CreateTable(const TableName: string): Boolean;
    function TableExists(const TableName: string): Boolean;
    function ValidateTableStructure(const TableName: string): Boolean;
    function DropTable(const TableName: string): Boolean;
    procedure ClearTable(const TableName: string);
    function EnsureTableReady(const TableName: string): Boolean;

    // DB Main Operations
    procedure LoadCSVtoDB(const CSVFile, TableName: string);
    procedure RemoveDuplicates(const TableName: string);
    procedure ExportToCSV(const TableName, CSVFile: string);
    
    // Additional funcs
    function StringToMoney(const S: string): Currency;
    function MoneyToString(const Value: Currency): string;
    function EscapeSQLString(const S: string): string;
    function GetFileSize(const FileName: string): Int64;
  public
    { Public declarations }
  end;

const
  INDEX_NAME = 'IX_Products_Brand_Article';
  EXPECTED_COLUMNS: array[0..3] of string = ('Brand', 'Article', 'Price', 'Quantity');
  EXPECTED_TYPES: array[0..3] of string = ('nvarchar', 'nvarchar', 'money', 'int');
  CONFIG_SECTION = 'Database';

var
  Form1: TForm1;

implementation

{$R *.dfm}

{ Загрузка конфигурации }

procedure TForm1.FormCreate(Sender: TObject);
begin
  FConfigFileName := ChangeFileExt(Application.ExeName, '.ini');
  LoadConfig;
end;

procedure TForm1.LoadConfig;
var
  Ini: TIniFile;
begin
  Ini := TIniFile.Create(FConfigFileName);
  try
    FConnectionString := Ini.ReadString(CONFIG_SECTION, 'ConnectionString', '');
    if FConnectionString = '' then
    begin
      CreateDefaultConfig;
      FConnectionString := Ini.ReadString(CONFIG_SECTION, 'ConnectionString', '');
    end;

    LogMessage('Конфигурация загружена из: ' + FConfigFileName);
  finally
    Ini.Free;
  end;
end;

procedure TForm1.CreateDefaultConfig;
var
  Ini: TIniFile;
  DefaultConnection: string;
begin
  Ini := TIniFile.Create(FConfigFileName);
  try
    DefaultConnection :=
      'Provider=SQLOLEDB.1;' +
      'Integrated Security=SSPI;' +
      'Initial Catalog=tempdb;' +
      'Data Source=(local);';
    
    Ini.WriteString(CONFIG_SECTION, 'ConnectionString', DefaultConnection);
    Ini.WriteString(CONFIG_SECTION, 'Description', 
      'Строка подключения к SQL Server. Для Windows-аутентификации используйте SSPI');
    
    LogMessage('Создан файл конфигурации по умолчанию: ' + FConfigFileName);
  finally
    Ini.Free;
  end;
end;

procedure TForm1.InitConnection;
begin
  FConnection := TADOConnection.Create(nil);
  FQuery := TADOQuery.Create(nil);

  FConnection.ConnectionString := FConnectionString;
  FConnection.LoginPrompt := False;
  
  try
    FConnection.Connected := True;
    FQuery.Connection := FConnection;
    LogMessage('Соединение с SQL Server установлено');
  except
    on E: Exception do
    begin
      LogMessage('ОШИБКА ПОДКЛЮЧЕНИЯ: ' + E.Message);
      LogMessage('Проверьте настройки в файле: ' + FConfigFileName);
      raise;
    end;
  end;
end;

procedure TForm1.CleanupConnection;
begin
  if Assigned(FQuery) then
    FreeAndNil(FQuery);
  if Assigned(FConnection) then
    FreeAndNil(FConnection);
end;

procedure TForm1.LogMessage(const Msg: string);
begin
  MemoLog.Lines.Add(FormatDateTime('hh:nn:ss', Now) + ': ' + Msg);
  Application.ProcessMessages;
end;

function TForm1.GetCSVFile: Boolean;
begin
  OpenDialog1.Filter := 'CSV файлы|*.csv|Все файлы|*.*';
  OpenDialog1.DefaultExt := 'csv';
  OpenDialog1.Title := 'Выберите CSV файл для загрузки';
  Result := OpenDialog1.Execute;
  if Result then
    FCSVFileName := OpenDialog1.FileName;
end;

procedure TForm1.btnCancelClick(Sender: TObject);
begin
  FCancelRequested := True;
  LogMessage('Отмена запрошена...');
end;

function TForm1.GetFileSize(const FileName: string): Int64;
var
  SR: TSearchRec;
begin
  if FindFirst(FileName, faAnyFile, SR) = 0 then
    Result := SR.Size
  else
    Result := 0;
  FindClose(SR);
end;

procedure TForm1.UpdateProgress(Current, Total: Int64);
var
  Percent: Integer;
begin
  if Total > 0 then
  begin
    Percent := Trunc((Current / Total) * 100);
    ProgressBar1.Position := Percent;
    ProgressBar1.Update;
  end;
end;

function TForm1.EscapeSQLString(const S: string): string;
begin
  Result := StringReplace(S, '''', '''''', [rfReplaceAll]);
end;

function TForm1.StringToMoney(const S: string): Currency;
var
  CleanedStr: string;
begin
  CleanedStr := Trim(S);
  CleanedStr := StringReplace(CleanedStr, ',', '.', []);
  CleanedStr := StringReplace(CleanedStr, ' ', '', [rfReplaceAll]);
  
  if not TryStrToCurr(CleanedStr, Result) then
    Result := 0;
end;

function TForm1.MoneyToString(const Value: Currency): string;
begin
  Result := FormatCurr('0.00', Value);
  Result := StringReplace(Result, '.', ',', []);
end;

procedure TForm1.DropIndexIfExists(const TableName, IndexName: string);
var
  IndexDropped: Integer;
begin
  FQuery.SQL.Text := Format(
    'DECLARE @dropped INT = 0; ' +
    'IF EXISTS (SELECT 1 FROM sys.indexes ' +
    '           WHERE name = ''%s'' AND object_id = OBJECT_ID(''%s'')) ' +
    'BEGIN ' +
    '  DROP INDEX %s ON %s; ' +
    '  SET @dropped = 1; ' +
    'END; ' +
    'SELECT @dropped as Dropped',
    [IndexName, TableName, IndexName, TableName]);
  
  FQuery.Open;
  IndexDropped := FQuery.FieldByName('Dropped').AsInteger;
  FQuery.Close;
  
  if IndexDropped = 1 then
    LogMessage(Format('Индекс %s удалён', [IndexName]));
end;

procedure TForm1.CreateIndex(const TableName, IndexName, Columns: string);
begin
  FQuery.SQL.Text := Format('CREATE INDEX %s ON %s(%s)', [IndexName, TableName, Columns]);
  FQuery.ExecSQL;
  LogMessage(Format('Индекс %s создан', [IndexName]));
end;

function TForm1.TableExists(const TableName: string): Boolean;
begin
  FQuery.SQL.Text := Format(
    'SELECT 1 FROM INFORMATION_SCHEMA.TABLES ' +
    'WHERE TABLE_NAME = ''%s'' AND TABLE_SCHEMA = ''dbo''',
    [TableName]);
  FQuery.Open;
  Result := not FQuery.Eof;
  FQuery.Close;
end;

function TForm1.ValidateTableStructure(const TableName: string): Boolean;
var
  i: Integer;
  DataType: string;
  AllValid: Boolean;
begin
  AllValid := True;
  FQuery.SQL.Text := Format(
    'SELECT COUNT(*) as Cnt FROM INFORMATION_SCHEMA.COLUMNS ' +
    'WHERE TABLE_NAME = ''%s''', [TableName]);

  FQuery.Open;
  if FQuery.FieldByName('Cnt').AsInteger <> 4 then
  begin
    FQuery.Close;
    LogMessage('Неверное количество столбцов');
    Result := False;
    Exit;
  end;
  FQuery.Close;

  for i := 0 to 3 do
  begin
    FQuery.SQL.Text := Format(
      'SELECT DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS ' +
      'WHERE TABLE_NAME = ''%s'' AND COLUMN_NAME = ''%s''',
      [TableName, EXPECTED_COLUMNS[i]]);
    FQuery.Open;
    
    if FQuery.Eof then
    begin
      LogMessage(Format('Столбец %s не найден', [EXPECTED_COLUMNS[i]]));
      AllValid := False;
    end
    else
    begin
      DataType := LowerCase(FQuery.FieldByName('DATA_TYPE').AsString);
      
      if (EXPECTED_TYPES[i] = 'nvarchar') and (Pos('char', DataType) = 0) then
      begin
        LogMessage(Format('Столбец %s: ожидается строковый тип, получен %s', 
          [EXPECTED_COLUMNS[i], DataType]));
        AllValid := False;
      end
      else if (EXPECTED_TYPES[i] = 'money') and (DataType <> 'money') and (DataType <> 'decimal') then
      begin
        LogMessage(Format('Столбец %s: ожидается money/decimal, получен %s', 
          [EXPECTED_COLUMNS[i], DataType]));
        AllValid := False;
      end
      else if (EXPECTED_TYPES[i] = 'int') and (DataType <> 'int') and (DataType <> 'integer') then
      begin
        LogMessage(Format('Столбец %s: ожидается целочисленный тип, получен %s', 
          [EXPECTED_COLUMNS[i], DataType]));
        AllValid := False;
      end;
    end;
    
    FQuery.Close;
    
    if not AllValid then
      Break;
  end;
  
  Result := AllValid;
  
  if Result then
    LogMessage('Структура таблицы корректна')
  else
    LogMessage('Структура таблицы не соответствует требуемой');
end;

function TForm1.CreateTable(const TableName: string): Boolean;
begin
  try
    FQuery.SQL.Text := Format(
      'CREATE TABLE %s (' +
      'ID INT IDENTITY(1,1) PRIMARY KEY, ' +
      'Brand NVARCHAR(255) NOT NULL, ' +
      'Article NVARCHAR(255) NOT NULL, ' +
      'Price MONEY NOT NULL, ' +
      'Quantity INT NOT NULL' +
      ')', [TableName]);
    FQuery.ExecSQL;
    LogMessage('Таблица создана');
    Result := True;
  except
    on E: Exception do
    begin
      LogMessage('Ошибка создания таблицы: ' + E.Message);
      Result := False;
    end;
  end;
end;

function TForm1.DropTable(const TableName: string): Boolean;
begin
  try
    FQuery.SQL.Text := Format('DROP TABLE %s', [TableName]);
    FQuery.ExecSQL;
    LogMessage('Таблица удалена');
    Result := True;
  except
    on E: Exception do
    begin
      LogMessage('Ошибка удаления таблицы: ' + E.Message);
      Result := False;
    end;
  end;
end;

procedure TForm1.ClearTable(const TableName: string);
begin
  try
    FQuery.SQL.Text := Format('TRUNCATE TABLE %s', [TableName]);
    FQuery.ExecSQL;
    LogMessage('Таблица очищена (TRUNCATE)');
  except
    try
      FQuery.SQL.Text := Format('DELETE FROM %s', [TableName]);
      FQuery.ExecSQL;
      LogMessage('Таблица очищена (DELETE)');
    except
      on E: Exception do
        LogMessage('Ошибка очистки таблицы: ' + E.Message);
    end;
  end;
end;

function TForm1.EnsureTableReady(const TableName: string): Boolean;
var
  UserChoice: Integer;
begin
  Result := False;
  
  if TableExists(TableName) then
  begin
    if ValidateTableStructure(TableName) then
    begin
      LogMessage('Таблица существует и имеет правильную структуру');
      
      UserChoice := MessageDlg(
        'Таблица уже существует. Выберите действие:' + #13#10 +
        #13#10 +
        'Да - использовать существующую таблицу (данные сохранятся)' + #13#10 +
        'Нет - очистить таблицу (удалить все записи)' + #13#10 +
        'Отмена - пересоздать таблицу (удалить и создать заново)',
        mtConfirmation, [mbYes, mbNo, mbCancel], 0);
      
      case UserChoice of
        mrYes: 
          begin
            LogMessage('Используется существующая таблица');
            Result := True;
          end;
          
        mrNo: 
          begin
            ClearTable(TableName);
            Result := True;
          end;
          
        mrCancel:
          begin
            LogMessage('Пересоздание таблицы...');
            if DropTable(TableName) then
              Result := CreateTable(TableName)
            else
              LogMessage('Не удалось пересоздать таблицу');
          end;
      end;
    end
    else
    begin
      UserChoice := MessageDlg(
        'Таблица существует, но имеет неверную структуру.' + #13#10 +
        'Пересоздать таблицу? (Нет - прервать выполнение)',
        mtWarning, [mbYes, mbNo], 0);
      
      if UserChoice = mrYes then
      begin
        if DropTable(TableName) then
          Result := CreateTable(TableName)
        else
          LogMessage('Не удалось пересоздать таблицу');
      end
      else
        LogMessage('Операция отменена пользователем из-за несоответствия структуры');
    end;
  end
  else
  begin
    LogMessage('Таблица не существует, будет создана новая');
    Result := CreateTable(TableName);
  end;
end;

procedure TForm1.LoadCSVtoDB(const CSVFile, TableName: string);
const
  BATCH_SIZE = 1000;
var
  F: TextFile;
  Line: string;
  LineNum, LoadedCount, BatchCounter: Integer;
  Fields, BatchValues: TStringList;
  FileSize, CurrentPos: Int64;
  FileOpened: Boolean;
  BaseSQL: string;
  ValuesSQL: string;
  i: Integer;
begin
  Fields := TStringList.Create;
  BatchValues := TStringList.Create;
  FileOpened := False;
  
  try
    FileSize := GetFileSize(CSVFile);
    CurrentPos := 0;
    
    AssignFile(F, CSVFile);
    Reset(F);
    FileOpened := True;
    LineNum := 0;
    LoadedCount := 0;
    BatchCounter := 0;
    
    LogMessage('Загрузка данных...');
    ProgressBar1.Position := 0;
    FCancelRequested := False;

    DropIndexIfExists(TableName, INDEX_NAME);

    BaseSQL := Format('INSERT INTO %s (Brand, Article, Price, Quantity) VALUES ',
      [TableName]);

    FConnection.Execute('BEGIN TRANSACTION');
    
    try
      while not Eof(F) and not FCancelRequested do
      begin
        ReadLn(F, Line);
        Inc(LineNum);
        CurrentPos := CurrentPos + Length(Line) + 2;
        
        if Trim(Line) = '' then
          Continue;

        Fields.Clear;
        while Line <> '' do
        begin
          if Pos(';', Line) > 0 then
          begin
            Fields.Add(Copy(Line, 1, Pos(';', Line) - 1));
            Delete(Line, 1, Pos(';', Line));
          end
          else
          begin
            Fields.Add(Line);
            Line := '';
          end;
        end;

        if Fields.Count >= 4 then
        begin
          Fields[0] := Trim(Fields[0]);  // Brand
          Fields[1] := Trim(Fields[1]);  // Article
          Fields[2] := Trim(Fields[2]);  // Price
          Fields[3] := Trim(Fields[3]);  // Quantity

          BatchValues.Add(Format('(N''%s'', N''%s'', %s, %d)',
            [
            EscapeSQLString(Fields[0]),
            EscapeSQLString(Fields[1]),
            StringReplace(FormatCurr('0.00', StringToMoney(Fields[2])), ',', '.', []),
            StrToIntDef(Fields[3], 0)
            ]));
          
          Inc(BatchCounter);
          Inc(LoadedCount);
          
          if (BatchCounter >= BATCH_SIZE) or Eof(F) then
          begin
            ValuesSQL := '';
            for i := 0 to BatchValues.Count - 1 do
            begin
              if i > 0 then
                ValuesSQL := ValuesSQL + ', ';
              ValuesSQL := ValuesSQL + BatchValues[i];
            end;

            FQuery.SQL.Text := BaseSQL + ValuesSQL;
            FQuery.ExecSQL;
            
            BatchValues.Clear;
            BatchCounter := 0;
            
            UpdateProgress(CurrentPos, FileSize);
            Application.ProcessMessages;
            
            if LoadedCount mod 10000 = 0 then
              LogMessage(Format('Загружено %d записей...', [LoadedCount]));
          end;
        end
        else
        begin
          LogMessage(Format('Предупреждение: строка %d содержит %d полей (ожидалось 4)',
            [LineNum, Fields.Count]));
        end;
      end;
      
      FConnection.Execute('COMMIT TRANSACTION');
      LogMessage(Format('Загружено записей: %d', [LoadedCount]));

      LogMessage('Создание индекса для оптимизации...');
      CreateIndex(TableName, INDEX_NAME, 'Brand, Article');
      
    except
      FConnection.Execute('ROLLBACK TRANSACTION');
      raise;
    end;
    
  finally
    if FileOpened then
      CloseFile(F);
    Fields.Free;
    BatchValues.Free;
    ProgressBar1.Position := 0;
  end;
end;

procedure TForm1.RemoveDuplicates(const TableName: string);
var
  BeforeCount, AfterCount: Integer;
  RowsDeleted: Integer;
begin
  FQuery.SQL.Text := Format('SELECT COUNT(*) as Cnt FROM %s', [TableName]);
  FQuery.Open;
  BeforeCount := FQuery.FieldByName('Cnt').AsInteger;
  FQuery.Close;

  FQuery.SQL.Text := Format(
    'WITH RankedProducts AS (' +
    '  SELECT ID, ROW_NUMBER() OVER (PARTITION BY Brand, Article ' +
    '    ORDER BY Price DESC, Quantity ASC, ID) AS RowNum ' +
    '  FROM %s' +
    ') DELETE FROM %s WHERE ID IN (SELECT ID FROM RankedProducts WHERE RowNum > 1)',
    [TableName, TableName]);
  
  FQuery.ExecSQL;

  RowsDeleted := FQuery.RowsAffected;
  AfterCount := BeforeCount - RowsDeleted;
  LogMessage(Format('Удалено дублей: %d', [RowsDeleted]));
  LogMessage(Format('Осталось записей: %d', [AfterCount]));
end;

procedure TForm1.ExportToCSV(const TableName, CSVFile: string);
const
  BUFFER_SIZE = 10000;
var
  F: TextFile;
  ExportedCount: Integer;
  Buffer: TStringList;
  FileOpened: Boolean;
begin
  FQuery.SQL.Text := Format('SELECT Brand, Article, Price, Quantity FROM %s ORDER BY Brand, Article', [TableName]);
  FQuery.Open;
  
  ExportedCount := 0;
  Buffer := TStringList.Create;
  FileOpened := False;
  
  try
    AssignFile(F, CSVFile);
    Rewrite(F);
    FileOpened := True;
    
    LogMessage('Выгрузка в файл...');
    ProgressBar1.Max := FQuery.RecordCount;
    ProgressBar1.Position := 0;
    
    while not FQuery.Eof and not FCancelRequested do
    begin
      Buffer.Add(Format('%s;%s;%s;%d',
        [
        FQuery.FieldByName('Brand').AsString,
        FQuery.FieldByName('Article').AsString,
        MoneyToString(FQuery.FieldByName('Price').AsCurrency),
        FQuery.FieldByName('Quantity').AsInteger
        ]));
      
      FQuery.Next;
      Inc(ExportedCount);
      
      if (Buffer.Count >= BUFFER_SIZE) or FQuery.Eof then
      begin
        Write(F, Buffer.Text);
        Buffer.Clear;
        
        ProgressBar1.Position := ExportedCount;
        Application.ProcessMessages;
        
        if ExportedCount mod 10000 = 0 then
          LogMessage(Format('Выгружено %d записей...', [ExportedCount]));
      end;
    end;
    
    LogMessage(Format('Выгружено записей: %d', [ExportedCount]));
    
  finally
    if FileOpened then
      CloseFile(F);
    FQuery.Close;
    Buffer.Free;
    ProgressBar1.Position := 0;
  end;
end;

{ MAIN }

procedure TForm1.btnProcessClick(Sender: TObject);
var
  OutputFileName: string;
begin
  if not GetCSVFile then
    Exit;

  MemoLog.Clear;
  LogMessage('=== НАЧАЛО ОБРАБОТКИ ===');
  LogMessage('Файл: ' + FCSVFileName);

  FTableName := '##TempProducts_' + FormatDateTime('hhnnss', Now);

  try
    InitConnection;

    if not EnsureTableReady(FTableName) then
    begin
      LogMessage('Операция прервана');
      Exit;
    end;
    
    LoadCSVtoDB(FCSVFileName, FTableName);
    
    if FCancelRequested then
    begin
      LogMessage('Операция отменена пользователем');
      Exit;
    end;
    
    RemoveDuplicates(FTableName);
    
    OutputFileName := ChangeFileExt(FCSVFileName, '_clean.csv');
    ExportToCSV(FTableName, OutputFileName);
    
    LogMessage('Результат: ' + OutputFileName);
    LogMessage('=== ОБРАБОТКА ЗАВЕРШЕНА ===');
    
    ShowMessage('Готово!');
    
  except
    on E: Exception do
    begin
      LogMessage('ОШИБКА: ' + E.Message);
      ShowMessage('Ошибка: ' + E.Message);
    end;
  end;
  
  CleanupConnection;
end;

end.

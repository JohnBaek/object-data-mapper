using System.Data.Common;
using System.Reflection;
using System.Text;
using FastExcel;
using Microsoft.Extensions.Configuration;
using ObjectDataMapper.Models;
using Oracle.ManagedDataAccess.Client;

namespace ObjectDataMapper.Providers; public class ObjectMapperProvider
{
    /// <summary>
    /// 설정파일
    /// </summary>
    private readonly IConfiguration Configuration;
    
    /// <summary>
    /// 생성자 
    /// </summary>
    /// <param name="configuration"></param>
    public ObjectMapperProvider(IConfiguration configuration)
    {
        Configuration = configuration;
    }
    
    /// <summary>
    /// 매핑을 시작한다.
    /// </summary>
    public async Task StartAsync()
    {
        Console.WriteLine("StartAsync . .");
        await using OracleConnection connection = new OracleConnection(Configuration.GetSection("ConnectionString").Value);

        try
        {
            connection.Open();
            Console.WriteLine("Connection Open . .");
            
            // 전체 테이블을 가져온다 
            List<ResponseTable> tables = await GatheringTables(connection);
            
            if (tables.Count == 0)
            {
                Console.WriteLine("Table Count 0");
                return;
            }
            
            // 테이블에 주석을 삽입한다.
            await AddComments(connection,tables);
            
            // 엑셀을 출력한다.
            string exportedExcelFilePath = ExportExcel(tables);
            
            Console.WriteLine($"Expoted Excel Completed : {exportedExcelFilePath}");

            await connection.CloseAsync();
            Console.WriteLine("Connection Close . .");
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
            throw ex;
        }
        finally
        {
            if(connection.KeepAlive)
                await connection.CloseAsync();
        }
    }

    /// <summary>
    /// 엑셀을 출력한다.
    /// </summary>
    /// <param name="tables"></param>
    private string ExportExcel(List<ResponseTable> tables)
    {
        try
        {
            // 파일정보를 초기화 한다.
            FileInfo fileInfo = InitializeFileInfo();
            
            // 엑셀 객체를 초기화한다.
            using FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(fileInfo);

            // 주석없는 모든 데이터 
            List<ResponseTable> withoutComments = tables.Where(i => !i.HasCommentInAllColumns).ToList();

            // 주석없는 모든 데이터 엑셀 파일작성
            writeInWorksheet(withoutComments,fastExcel,1);
            
            // 주석있는 모든 데이터 
            List<ResponseTable> withComments = tables.Where(i => i.HasCommentInAllColumns).ToList();
            
            // 주석있는 모든 데이터 엑셀 파일작성
            writeInWorksheet(withComments,fastExcel,2);

            return fileInfo.FullName;
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }
    }

    /// <summary>
    /// 워크시트 작성
    /// </summary>
    /// <param name="tables"></param>
    /// <param name="fastExcel"></param>
    /// <param name="sheetNumber"></param>
    private void writeInWorksheet(List<ResponseTable> tables,  FastExcel.FastExcel fastExcel, int sheetNumber)
    {
        try
        {
            List<Row> rows;
            
            // 기본 워크시트 생성
            Worksheet sheet = GetDefaultWorksheet(out rows);
            

            int currentRow = 1;
            
            // 전체 처리
            foreach (ResponseTable table in tables)
            {
                // 테이블명 삽입
                rows.Add(new Row(++currentRow , new List<Cell>(){new Cell(1, table.Name)}) );

                
                // 모든컬럼에 대해 처리 
                foreach (ResponseColumn column in table.columns)
                {
                    List<Cell> cells = new List<Cell>();
                    // 컬럼명
                    cells.Add(new Cell(2,column.Name));
                    
                    // 주석이 없는 경우 
                    if(String.IsNullOrWhiteSpace(column.Comments))
                        cells.Add(new Cell(3,"-"));
                    // 주석이 있는경우 
                    else
                        cells.Add(new Cell(3,column.Comments));
                    
                    // 한줄 추가
                    rows.Add(new Row(currentRow++ , cells ));
                }
            }
   
            // 업데이트
            fastExcel.Update(sheet,sheetNumber);
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }
    }

    /// <summary>
    /// 헤더를 포함한 기본 워크시트를 리턴한다.
    /// </summary>
    /// <returns></returns>
    private Worksheet GetDefaultWorksheet(out List<Row> rows)
    {
        // 처리할 첫번재 워크시트 : 주석이 달리지않은 컬럼
        Worksheet newWorkSheet = new Worksheet();
        
        try
        {
            // 로우 설정 ( 헤더 )
            rows = new List<Row>();
            
            // 셀 설정 ( 헤더 )
            List<Cell> headerCells = new List<Cell>();
            
            // 헤더추가 
            headerCells.Add(new Cell(1, "테이블명"));
            headerCells.Add(new Cell(2, "컬럼명"));
            headerCells.Add(new Cell(3, "주석"));
            
            // 헤더로우 추가
            rows.Add(new Row(1,headerCells));

            newWorkSheet.Rows = rows;
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }

        return newWorkSheet;
    }

    /// <summary>
    /// 파일 및 디렉토리를 초기화하고 목적지 파일을 리턴한다.
    /// </summary>
    /// <returns></returns>
    private FileInfo InitializeFileInfo()
    {
        FileInfo fileInfo = null;
        try
        {
            // 디렉토리 루트 
            string directoryPath = System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

            // 출력될 디렉토리 
            string exportedDirectory = Path.Combine(directoryPath, "Exported");

            // 출력 디렉토리 정보를 가져온다 
            DirectoryInfo directoryInfo = new DirectoryInfo(exportedDirectory);
            
            // 디렉토리가 존재하지않는경우 
            if(!directoryInfo.Exists)
                // 디렉토리 생성
                directoryInfo.Create();
            
            // 복사할 파일 
            string sourceFilePath = Path.Combine(directoryPath, "Resources", "result.xlsx");
            
            // 복사할 파일 객체
            FileInfo sourceFileInfo = new FileInfo(sourceFilePath);

            // 복사될 파일 경로를 설정한다.
            string destFilePath = Path.Combine(exportedDirectory, $"{Guid.NewGuid().ToString()}.xlsx" );

            // 파일 복사
            sourceFileInfo.CopyTo(destFilePath);

            return new FileInfo(destFilePath);
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }

        return fileInfo;
    }

    /// <summary>
    /// 주석을 삽입한다.
    /// </summary>
    /// <param name="connection"></param>
    /// <param name="tables"></param>
    /// <exception cref="NotImplementedException"></exception>
    private async Task AddComments(OracleConnection connection, List<ResponseTable> tables)
    {
        Console.WriteLine("AddComments . .");
        StringBuilder query = new StringBuilder();
        int processCount = 0;
        try
        {
            // 모든테이블에 대해 처리한다.
            foreach (ResponseTable table in tables)
            {
                Console.WriteLine($"[{++processCount}/{tables.Count}] TABLE : { table.Name }");
                
                // 커맨드 객체를 만든다.
                OracleCommand oracleCommand = new OracleCommand();
                oracleCommand.Connection = connection;

                // 테이블 조회 Statement 작성
                query.Clear();
                query.AppendLine($"SELECT OWNER, TABLE_NAME, Column_name, COMMENTS");
                query.AppendLine($"From ALL_COL_COMMENTS");
                query.AppendLine($"where TABLE_NAME = '{table.Name}'");
                oracleCommand.CommandText = query.ToString();
                
                // 쿼리를 실행한다.
                DbDataReader reader = await oracleCommand.ExecuteReaderAsync();
                
                // 삽입할 객체
                List<ResponseColumn> columns = table.columns;

                // 모든 객체가 주석을 갖고있는지 여부
                bool isHaveAllComments = true;
                
                // 레코드를 가져온다
                while (await reader.ReadAsync())
                {
                    // 테이블 명을 읽어온다. 
                    string columnName = reader["COLUMN_NAME"] as string;
                    string comments = reader["COMMENTS"] as string;
                    columns.Add(new ResponseColumn()
                    {
                        Name = columnName,
                        Comments = comments
                    });

                    // 주석이 비어있는 경우 
                    if (String.IsNullOrWhiteSpace(comments))
                        isHaveAllComments = false;
                }

                table.HasCommentInAllColumns = isHaveAllComments;
            }

            Console.WriteLine($"컬럼에 전체 주석이 존재하는 테이블 수: { tables.Count(i => i.HasCommentInAllColumns) }");
            Console.WriteLine($"컬럼에 전체 주석이 하나라도 없는 테이블 수: { tables.Count(i => !i.HasCommentInAllColumns) }");
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }
    }

    /// <summary>
    /// 테이블 정보를 수집한다.
    /// </summary>
    /// <param name="connection"></param>
    /// <returns></returns>
    private async Task<List<ResponseTable>> GatheringTables(OracleConnection connection)
    {
        Console.WriteLine("GatheringTables . .");
        List<ResponseTable> result = new List<ResponseTable>();
        StringBuilder query = new StringBuilder();
        
        try
        {
            // 커맨드 객체를 만든다.
            OracleCommand oracleCommand = new OracleCommand();
            oracleCommand.Connection = connection;

            // 테이블 조회 Statement 작성
            query.Clear();
            query.AppendLine("SELECT TABLE_NAME");
            query.AppendLine("FROM all_tables");
            query.AppendLine("where OWNER = 'HANDLE4UDEV'");
            oracleCommand.CommandText = query.ToString();

            // 쿼리를 실행한다.
            DbDataReader reader = await oracleCommand.ExecuteReaderAsync();
            
            // 레코드를 가져온다
            while (await reader.ReadAsync())
            {
                // 테이블 명을 읽어온다. 
                string tableName = reader["TABLE_NAME"] as string;
                result.Add(new ResponseTable() {Name = tableName}); 
            }
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            throw;
        }
        return result;
    }
}
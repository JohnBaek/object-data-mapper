using Microsoft.Extensions.Configuration;
using Oracle.ManagedDataAccess.Client;

namespace ObjectDataMapper.Providers;

public class ObjectMapperProvider
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
}
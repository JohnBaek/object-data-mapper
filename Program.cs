using Microsoft.Extensions.Configuration;
using ObjectDataMapper.Providers;

// 빌더를 생성한다.
IConfigurationBuilder builder = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: false);

// 빌더를 통해 컨피그 파일을 가져온다. 
IConfiguration config = builder.Build();

// 프로바이더를 생성한다.
ObjectMapperProvider objectMapperProvider = new ObjectMapperProvider(config);

// 프로바이더 시작
await objectMapperProvider.StartAsync();
using Ichiba.Libs.DocumentSdk.Abstractions;
using Ichiba.Libs.DocumentSdk.Connectors;
using Ichiba.Libs.DocumentSdk.Constants;
using Ichiba.Libs.DocumentSdk.Models;
using Ichiba.Libs.DocumentSdk.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Polly;
using RestEase.HttpClientFactory;

namespace Ichiba.Libs.DocumentSdk;

public static class ConfigureServices
{
    public static IServiceCollection AddDocumentSdk(
        this IServiceCollection services,
        IConfiguration configuration,
        Action<IHttpClientBuilder>? configHttpClientBuilder = null,
        bool ignoreSslCertValidate = false,
        bool useForwardToken = false)
    {
        services.AddTransient<IPdfService, PdfService>();
        services.AddTransient<IWordService, WordService>();
        services.AddScoped(typeof(IExcelService<>), typeof(ExcelService<>));
        services.AddScoped<IDocumentServiceFactory, DocumentServiceFactory>();
        RegisterDocumentValidator(services);
        var license = new Aspose.Cells.License();
        license.SetLicense(AsposeLicenseConstants.LicenseCell);
        // Get Config
        var cfgSection = configuration.GetSection("Bff");
        var cfg = cfgSection.Get<BffSettings>();

        return services.AddDocumentSdk(
            cfg!.BaseUrl,
            configHttpClientBuilder,
            ignoreSslCertValidate,
            useForwardToken);
    }

    private static void RegisterDocumentValidator(IServiceCollection services)
    {
        var documentValidatorType = typeof(IDocumentValidator<>);

        var types = AppDomain.CurrentDomain.GetAssemblies()
            .SelectMany(a => a.GetTypes())
            .Where(t => t.IsClass
                        && !t.IsAbstract
                        && t.GetInterfaces().Any(i =>
                            i.IsGenericType &&
                            i.GetGenericTypeDefinition() == documentValidatorType));

        foreach (var type in types)
        {
            var interfaceType = type.GetInterfaces()
                .First(i => i.IsGenericType &&
                            i.GetGenericTypeDefinition() == documentValidatorType);

            var genericArgument = interfaceType.GetGenericArguments()[0];

            if (!typeof(DocumentItemBase).IsAssignableFrom(genericArgument))
            {
                continue;
                // throw new ArgumentException($"The type {genericArgument.Name} in {type.Name} does not inherit from DocumentItemBase.");
            }

            services.AddTransient(interfaceType, type);
        }
    }

    private static IServiceCollection AddDocumentSdk(
        this IServiceCollection services,
        string baseUrl,
        Action<IHttpClientBuilder>? configHttpClientBuilder = null,
        bool ignoreSslCertValidate = false,
        bool useForwardToken = false)
    {
        var httpClientBuilder = services
            .AddHttpClient(typeof(ConfigureServices).Namespace)
            .ConfigureHttpClient((_, x) => { x.BaseAddress = new Uri(baseUrl); });

        if (ignoreSslCertValidate)
        {
            httpClientBuilder = httpClientBuilder.ConfigurePrimaryHttpMessageHandler(
                () => new HttpClientHandler
                {
                    ServerCertificateCustomValidationCallback =
                        (_, _, _, _) => true
                });
        }

        httpClientBuilder = ConfigHttpServices(httpClientBuilder);

        // other config: HttpMessageDelegateHandler , custom Response result handler.........
        configHttpClientBuilder?.Invoke(httpClientBuilder);

        // Retry Policy
        ConfigureRetryPolicy(httpClientBuilder);
        // if (useForwardToken)
        //     httpClientBuilder.AddHttpMessageHandler<ForwardTokenHttpMessageDelegateHandler>();
        return services;
    }

    /// <summary>
    /// Config  Retry Policy
    /// </summary>
    private static IHttpClientBuilder ConfigureRetryPolicy(IHttpClientBuilder httpClientBuilder)
    {
        httpClientBuilder
            .AddTransientHttpErrorPolicy(builder =>
                builder.WaitAndRetryAsync(new[]
                {
                    TimeSpan.FromMilliseconds(100),
                    TimeSpan.FromMilliseconds(500),
                    TimeSpan.FromSeconds(2)
                }));
        return httpClientBuilder;
    }

    /// <summary>
    /// Register Http Services
    /// </summary>
    private static IHttpClientBuilder ConfigHttpServices(IHttpClientBuilder httpClientBuilder)
    {
        #region Register API Service

        httpClientBuilder.UseWithRestEaseClient<IDocumentConnector>();

        #endregion Register API Service

        return httpClientBuilder;
    }
}

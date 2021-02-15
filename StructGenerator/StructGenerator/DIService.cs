using Microsoft.Extensions.DependencyInjection;
using StructGenerator.GeneratorFactory;

namespace StructGenerator
{
    public class DIService
    {
        private readonly IServiceCollection _service;

        public DIService(IServiceCollection service)
        {
            _service = service;

        }
        public IServiceCollection Inject()
        {
            _service.AddSingleton<AppSetting>();
            _service.AddSingleton<Generator>();
            _service.AddSingleton<IFile, StructFormat>();
            return _service;
        }
     }
}

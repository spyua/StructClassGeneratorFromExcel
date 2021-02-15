using Microsoft.Extensions.DependencyInjection;
using StructGenerator.GeneratorFactory;
using System;

namespace StructGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            var sysDI = new DIService(new ServiceCollection());
            var serviceContainer = sysDI.Inject();
            var provider = serviceContainer.BuildServiceProvider();

            var generator = provider.GetService<Generator>();
            generator.GenStruct();

            Console.ReadKey();


        }
    }
}

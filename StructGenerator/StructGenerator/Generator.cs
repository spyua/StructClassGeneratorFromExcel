using StructGenerator.GeneratorFactory;
using System;
using System.Collections.Generic;
using System.Text;

namespace StructGenerator
{
    public class Generator
    {

        private IFile _file;

        public Generator(IFile file)
        {
            _file = file;
        }

        public void GenStruct()
        {

            _file.DecodeExcel();
            _file.GenFile();
        }
    }
}

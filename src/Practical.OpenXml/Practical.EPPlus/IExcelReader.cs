using System.IO;

namespace Practical.EPPlus
{
    public interface IExcelReader<T>
    {
        T Read(Stream stream);
    }
}

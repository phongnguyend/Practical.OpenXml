using System.IO;

namespace Practical.EPPlus
{
    public interface IExcelWriter<T>
    {
        void Write(T data, Stream stream);
    }
}

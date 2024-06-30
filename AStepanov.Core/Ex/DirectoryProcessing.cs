using System.Reflection;

namespace AStepanov.Core.Ex
{
    public static class DirectoryProcessing
    {
        public static string Directory(this Assembly assembly)
        {
            return Path.GetDirectoryName(assembly.Location);
        }
    }
}

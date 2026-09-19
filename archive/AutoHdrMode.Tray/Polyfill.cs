// Polyfill: records need IsExternalInit, which .NET Framework 4.7.2 lacks.
namespace System.Runtime.CompilerServices
{
    internal static class IsExternalInit { }
}

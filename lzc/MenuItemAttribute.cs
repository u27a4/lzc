using System;
using System.ComponentModel.Composition;

namespace lzc
{
    [AttributeUsage(AttributeTargets.Method)]
    public sealed class MenuItemAttribute : ExportAttribute { }
}

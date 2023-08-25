// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.



#if NETFRAMEWORK
using System.ComponentModel;
// Enable c#9.0 record class width .net framework target; see https://sergiopedri.medium.com/enabling-and-using-c-9-features-on-older-and-unsupported-runtimes-ce384d8debb
namespace System.Runtime.CompilerServices {
    /// <summary>
    /// Reserved to be used by the compiler for tracking metadata.
    /// This class should not be used by developers in source code.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    internal static class IsExternalInit {
    }
}
namespace System.Runtime.CompilerServices {
    [AttributeUsage(AttributeTargets.Parameter, AllowMultiple = false, Inherited = false)]
    internal sealed class CallerArgumentExpressionAttribute : Attribute {
        public CallerArgumentExpressionAttribute(string parameterName) {
            ParameterName = parameterName;
        }

        public string ParameterName { get; }
    }
}
namespace System.Diagnostics.CodeAnalysis {
    /// <summary>
    /// Specifies that an output will not be null even if the corresponding type allows it. Specifies that an input argument was not null when the call returns.
    /// </summary>
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Parameter | AttributeTargets.Property | AttributeTargets.ReturnValue, Inherited = false)]
    internal sealed class NotNullAttribute : Attribute {
    }
    /// <summary>
    /// Applied to a method that will never return under any circumstance.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, Inherited = false)]
    internal sealed class DoesNotReturnAttribute : Attribute {
    }
}
#endif

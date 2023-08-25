using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Bs.XML.SpreadSheet {
    internal static class Throw {
        [DoesNotReturn] 
        internal static void InvalidOperationException() {
            throw new InvalidOperationException();
        }
        [DoesNotReturn] 
        internal static void InvalidOperationException(FormattableString message) {
            throw new InvalidOperationException(message.ToString());
        }
        internal static void IfNull([NotNull]object? argument, FormattableString errMessage, [CallerArgumentExpression("argument")] string? paramName = null) {
            if (argument is null) {
                throw new ArgumentNullException(paramName, errMessage.ToString());
            }
        }
        internal static void IfNull([NotNull]object? argument, Func<FormattableString> getErrMessage, [CallerArgumentExpression("argument")] string? paramName = null) {
            if (argument is null) {
                throw new ArgumentNullException(paramName, getErrMessage().ToString());
            }
        }
        /// <summary>Throws an <see cref="ArgumentNullException"/> if <paramref name="argument"/> is null.</summary>
        /// <param name="argument">The reference type argument to validate as non-null.</param>
        /// <param name="paramName">The name of the parameter with which <paramref name="argument"/> corresponds.</param>
        internal static void IfNull([NotNull] object? argument, string? message=null, [CallerArgumentExpression("argument")] string? paramName = null) {
            if (argument is null) {
                throw new ArgumentNullException(paramName, message);
            }
        }
        internal static void IfNull([NotNull] object? argument, Func<Exception> throwEx) {
            if (argument is null) {
                throw throwEx();
            }
        }
        internal static void IfIsTrue(bool argument, Func<Exception> throwEx) {
            if (argument) {
                throw throwEx();
            }
        }
        
    }
}
using System;
using System.Management.Automation;

namespace PSWindowsUpdate {
    internal class ValidateDateTimeAttribute : ValidateArgumentsAttribute {
        protected override void Validate(object arguments, EngineIntrinsics engineIntrinsics) {
            if ((DateTime)arguments < DateTime.Now) {
                throw new SystemException("Execution time is gone.");
            }
        }
    }
}
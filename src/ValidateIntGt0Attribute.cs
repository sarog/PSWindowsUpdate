using System;
using System.Management.Automation;

namespace PSWindowsUpdate {
    internal class ValidateIntGt0Attribute : ValidateArgumentsAttribute {
        protected override void Validate(object arguments, EngineIntrinsics engineIntrinsics) {
            if ((int)arguments <= 0) {
                throw new SystemException("Value must be greater than 0.");
            }
        }
    }
}
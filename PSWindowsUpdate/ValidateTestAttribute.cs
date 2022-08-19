using System.Management.Automation;

namespace PSWindowsUpdate {
    internal class ValidateTestAttribute : ValidateArgumentsAttribute {
        protected override void Validate(object arguments, EngineIntrinsics engineIntrinsics) { }
    }
}
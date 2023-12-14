using System.Management.Automation;

namespace PSWindowsUpdate {
    /// <summary>ValidateTestAttribute</summary>
    internal class ValidateTestAttribute : ValidateArgumentsAttribute {
        protected override void Validate(object arguments, EngineIntrinsics engineIntrinsics) { }
    }
}
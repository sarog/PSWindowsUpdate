using System;
using System.Management.Automation;

namespace PSWindowsUpdate
{
    /// <summary>ValidateRecurseCycleAttribute</summary>
    internal class ValidateRecurseCycleAttribute : ValidateArgumentsAttribute
    {
        protected override void Validate(object arguments, EngineIntrinsics engineIntrinsics)
        {
            if ((int)arguments <= 1)
            {
                throw new SystemException("Recursive cycle must be greater than 1. First run is the main cycle.");
            }
        }
    }
}
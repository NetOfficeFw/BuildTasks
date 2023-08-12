using System;
using System.Collections;
using Microsoft.Build.Framework;
using NUnit.Framework;

namespace NetOfficeFw.Build
{
    class BuildEngineStub : IBuildEngine
    {
        public bool ContinueOnError => true;

        public int LineNumberOfTaskNode => 0;

        public int ColumnNumberOfTaskNode => 0;

        public string ProjectFileOfTaskNode => "UnitTest.proj";

        public bool BuildProjectFile(string projectFileName, string[] targetNames, IDictionary globalProperties, IDictionary targetOutputs)
        {
            return false;
        }

        public void LogCustomEvent(CustomBuildEventArgs e)
        {
        }

        public void LogErrorEvent(BuildErrorEventArgs e)
        {
            Assert.Fail($"Build task failed. Code={e.Code}, Message={e.Message}");
        }

        public void LogMessageEvent(BuildMessageEventArgs e)
        {
        }

        public void LogWarningEvent(BuildWarningEventArgs e)
        {
        }
    }
}

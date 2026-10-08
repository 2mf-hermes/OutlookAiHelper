using System;
using OutlookAiHelper.Tests;

namespace OutlookAiHelper.TestHost
{
    public static class TestRunner
    {
        public static int Main(string[] args)
        {
            var failed = 0;
            failed += RuleEngineTests.Run();
            failed += TodoStoreTests.Run();
            failed += SettingsPrivacyTests.Run();
            failed += ScanScopeTests.Run();
            failed += AutoRefreshTests.Run();
            failed += MailThreadTests.Run();
            failed += MailOrderingTests.Run();
            failed += UpdateTests.Run();
            if (failed == 0)
            {
                Console.WriteLine("ALL TESTS PASSED");
                return 0;
            }

            Console.WriteLine("TESTS FAILED: " + failed);
            return 1;
        }
    }
}

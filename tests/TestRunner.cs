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
            failed += MailSourceTests.Run();
            failed += GraphMailTests.Run();
            failed += GraphJsonTests.Run();
            failed += GraphAuthTests.Run();
            failed += RefreshTokenStoreTests.Run();
            failed += GraphReaderTests.Run();
            failed += GraphSignInTests.Run();
            failed += SignInReasonTests.Run();
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

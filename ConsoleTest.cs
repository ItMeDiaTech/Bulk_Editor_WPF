using System;
using BulkEditor.Core.Configuration;
using BulkEditor.Infrastructure.Services;

class Program
{
    static void Main(string[] args)
    {
        try
        {
            Console.WriteLine("Testing core dependencies...");

            // Test 1: Basic configuration classes
            var appSettings = new AppSettings();
            Console.WriteLine("✓ AppSettings created successfully");

            // Test 2: Create SerilogService
            var logger = new BulkEditor.Infrastructure.Services.SerilogService();
            Console.WriteLine("✓ SerilogService created successfully");

            // Test 3: Create ConfigurationService
            var configService = new ConfigurationService(logger);
            Console.WriteLine("✓ ConfigurationService created successfully");

            Console.WriteLine("All core dependencies work correctly!");
            Console.WriteLine("The issue is likely in WPF startup or MaterialDesign themes.");

        }
        catch (Exception ex)
        {
            Console.WriteLine($"ERROR: {ex.Message}");
            Console.WriteLine($"Stack Trace: {ex.StackTrace}");
        }

        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}
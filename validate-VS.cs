#:package Microsoft.VisualStudio.Setup.Configuration.Interop@3.10.2154
#:property BuiltInComInteropSupport=true
#:property EnableComHosting=true

using Microsoft.VisualStudio.Setup.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

// Get the setup configuration
var query = GetSetupConfiguration();
if (query == null)
{
    Console.WriteLine("No Visual Studio installations found.");
    return;
}

// Find all VS instances
var instances = GetVisualStudioInstances(query);
if (!instances.Any())
{
    Console.WriteLine("No Visual Studio instances found.");
    return;
}

// List available instances
Console.WriteLine("Available Visual Studio installations:");
Console.WriteLine();

for (int i = 0; i < instances.Count; i++)
{
    var instance = instances[i];
    Console.WriteLine($"{i + 1}. {GetFriendlyName(instance)}");
    Console.WriteLine($"   Version: {instance.GetInstallationVersion()}");
    Console.WriteLine($"   Path: {instance.GetInstallationPath()}");
    Console.WriteLine();
}

// Get user selection
Console.Write($"Select an installation (1-{instances.Count}): ");
var input = Console.ReadLine();

if (!int.TryParse(input, out int selection) || selection < 1 || selection > instances.Count)
{
    Console.WriteLine("Invalid selection.");
    return;
}

// Display information about selected instance
var selectedInstance = instances[selection - 1];
Console.WriteLine();
Console.WriteLine("=== Selected Visual Studio Installation ===");
Console.WriteLine($"Display Name: {GetFriendlyName(selectedInstance)}");
Console.WriteLine($"Installation Path: {selectedInstance.GetInstallationPath()}");
Console.WriteLine($"Version: {selectedInstance.GetInstallationVersion()}");
Console.WriteLine($"Instance ID: {selectedInstance.GetInstanceId()}");

// Helper methods
static ISetupConfiguration2? GetSetupConfiguration()
{
    try
    {
        return new SetupConfiguration() as ISetupConfiguration2;
    }
    catch (COMException ex) when (ex.HResult == unchecked((int)0x80040154)) // REGDB_E_CLASSNOTREG
    {
        // VS setup API not available
        return null;
    }
}

static List<ISetupInstance2> GetVisualStudioInstances(ISetupConfiguration2 query)
{
    var instances = new List<ISetupInstance2>();
    
    try
    {
        var enumInstances = query.EnumAllInstances();
        var instanceArray = new ISetupInstance[1];
        
        while (true)
        {
            enumInstances.Next(1, instanceArray, out int fetched);
            if (fetched == 0) break;
            
            if (instanceArray[0] is ISetupInstance2 instance2)
            {
                // Only include instances that are installed and have the VS product
                var state = instance2.GetState();
                if ((state & InstanceState.Local) == InstanceState.Local)
                {
                    instances.Add(instance2);
                }
            }
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error enumerating instances: {ex.Message}");
    }
    
    return instances;
}

static string GetFriendlyName(ISetupInstance2 instance)
{
    try
    {
        var displayName = instance.GetDisplayName();
        if (!string.IsNullOrEmpty(displayName))
            return displayName;
        
        var product = instance.GetProduct();
        if (!string.IsNullOrEmpty(product?.GetId()))
        {
            var productId = product.GetId();
            return productId switch
            {
                "Microsoft.VisualStudio.Product.Enterprise" => "Visual Studio Enterprise",
                "Microsoft.VisualStudio.Product.Professional" => "Visual Studio Professional",
                "Microsoft.VisualStudio.Product.Community" => "Visual Studio Community",
                "Microsoft.VisualStudio.Product.BuildTools" => "Visual Studio Build Tools",
                _ => "Visual Studio"
            };
        }
        
        return "Visual Studio";
    }
    catch
    {
        return "Visual Studio (Unknown)";
    }
}
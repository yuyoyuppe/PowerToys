// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.Json.Serialization;
using Microsoft.PowerToys.Settings.UI.Library;

namespace PowerToys.Settings.DSC.Schema
{
    internal sealed class Program
    {
        private sealed class ModuleSettingsStructure
        {
            public string ModuleName { get; set; }

            public Type PropertiesType { get; set; }
        }

        private static void ListPropertiesAndAttributes(ModuleSettingsStructure moduleSettings)
        {
            foreach (PropertyInfo property in moduleSettings.PropertiesType.GetProperties())
            {
                string propertyInfo = $"{property.Name} ({property.PropertyType.Name})";

                var jsonIgnoreAttr = property.GetCustomAttribute<JsonIgnoreAttribute>();
                var jsonConverterAttr = property.GetCustomAttribute<JsonConverterAttribute>();
                var jsonPropertyNameAttr = property.GetCustomAttribute<JsonPropertyNameAttribute>();

                if (jsonIgnoreAttr != null)
                {
                    propertyInfo += " [JsonIgnore]";
                }

                if (jsonConverterAttr != null && jsonConverterAttr.ConverterType == typeof(BoolPropertyJsonConverter))
                {
                    propertyInfo += " [JsonConverter(BoolPropertyJsonConverter)]";
                }

                if (jsonPropertyNameAttr != null)
                {
                    propertyInfo += $" [JsonPropertyName(\"{jsonPropertyNameAttr.Name}\")]";
                }

                Console.WriteLine(propertyInfo);
            }

            Console.WriteLine(string.Empty);
        }

        public static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to a .NET DLL.");
                return;
            }

            string dllPath = args[0];
            try
            {
                Assembly assembly = Assembly.LoadFrom(dllPath);
                Type[] types = assembly.GetTypes();

                var moduleSettings = types.Where(type => type.IsClass && type.FullName.EndsWith("Settings", StringComparison.InvariantCulture)).Select(type =>
                {
                    var propertiesInfo = type.GetProperty("Properties");
                    var moduleNameField = type.GetField("ModuleName", BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy);

                    if (propertiesInfo != null && propertiesInfo.PropertyType.IsClass && moduleNameField != null && moduleNameField.IsLiteral && !moduleNameField.IsInitOnly && moduleNameField.FieldType == typeof(string))
                    {
                        var moduleName = moduleNameField.GetValue(null).ToString();
                        return new ModuleSettingsStructure { ModuleName = moduleName, PropertiesType = propertiesInfo.PropertyType };
                    }
                    else
                    {
                        return null;
                    }
                }).Where(settings => settings != null).ToList();

                foreach (var settings in moduleSettings)
                {
                    Console.WriteLine($"{settings.ModuleName}: {settings.PropertiesType}");
                    ListPropertiesAndAttributes(settings);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}

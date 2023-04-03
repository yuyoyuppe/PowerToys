// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Diagnostics;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace MouseWithoutBordersService
{
    public delegate void StopService();

    public static class CmdArgs
    {
        public static string[] Value { get; set; }

        public static StopService StopServiceDelegate { get; set; }
    }

    internal sealed class Program
    {
        [STAThread]
        private static void Main()
        {
            string[] args = Environment.GetCommandLineArgs();
            CmdArgs.Value = args;
            var builder = Host.CreateDefaultBuilder(args);

            var host = builder
            .UseWindowsService(options =>
            {
                options.ServiceName = "Mouse Without Borders service";
            })
            .ConfigureServices(services =>
            {
                services.AddHostedService<Worker>();
            })
            .Build();

            CmdArgs.StopServiceDelegate = async () => { await host.StopAsync(); };
            host.Run();
            host.StopAsync();
        }
    }
}

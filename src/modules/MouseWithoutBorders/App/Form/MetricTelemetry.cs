// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

// <summary>
//     Startup/main form + helper methods.
// </summary>
// <history>
//     2008 created by Truong Do (ductdo).
//     2009-... modified by Truong Do (TruongDo).
// </history>
namespace MouseWithoutBorders.Form
{
    internal class MetricTelemetry
    {
        private readonly string v;
        private readonly int switchCount;
        private readonly double value;

        public MetricTelemetry(string v, int switchCount)
        {
            this.v = v;
            this.switchCount = switchCount;
        }

        public MetricTelemetry(string v, double value)
        {
            this.v = v;
            this.value = value;
        }
    }
}

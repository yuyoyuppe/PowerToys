// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Collections.Generic;

// <summary>
//     Keyboard/Mouse hook callback implementation.
// </summary>
// <history>
//     2008 created by Truong Do (ductdo).
//     2009-... modified by Truong Do (TruongDo).
// </history>
using MouseWithoutBorders.Form;

namespace MouseWithoutBorders.Class
{
    internal class TelemetryClient
    {
        internal void Flush()
        {
            throw new NotImplementedException();
        }

        internal void TrackEvent(string v)
        {
            throw new NotImplementedException();
        }

        internal void TrackEvent(string v, Dictionary<string, string> dictionary)
        {
            throw new NotImplementedException();
        }

        internal void TrackException(Exception e, Dictionary<string, string> dictionary)
        {
            throw new NotImplementedException();
        }

        internal void TrackException(Exception e)
        {
            throw new NotImplementedException();
        }

        internal void TrackMetric(MetricTelemetry metricTelemetry)
        {
            throw new NotImplementedException();
        }

        internal void TrackTrace(string log, SeverityLevel severityLevel, Dictionary<string, string> dictionary)
        {
            throw new NotImplementedException();
        }

        internal void TrackTrace(string v)
        {
            throw new NotImplementedException();
        }
    }
}

#include "pch.h"
#include "trace.h"

TRACELOGGING_DEFINE_PROVIDER(
  g_hProvider,
  "Microsoft.PowerToys",
  // {38e8889b-9731-53f5-e901-e8a7c1753074}
  (0x38e8889b, 0x9731, 0x53f5, 0xe9, 0x01, 0xe8, 0xa7, 0xc1, 0x75, 0x30, 0x74),
  TraceLoggingOptionProjectTelemetry());

void Trace::RegisterProvider() {
  TraceLoggingRegister(g_hProvider);
}

void Trace::UnregisterProvider() {
  TraceLoggingUnregister(g_hProvider);
}

void Trace::EventShow() {
  TraceLoggingWrite(
    g_hProvider,
    "MTND::Event::ShowOverlay",
    ProjectTelemetryPrivacyDataTag(ProjectTelemetryTag_ProductAndServicePerformance),
    TraceLoggingBoolean(TRUE, "UTCReplace_AppSessionGuid"),
    TraceLoggingKeyword(PROJECT_KEYWORD_MEASURE));
}

void Trace::EventHide(const __int64 duration_ms) {
  TraceLoggingWrite(
    g_hProvider,
    "MTND::Event::HideOverlay",
    TraceLoggingInt64(duration_ms, "Duration in ms"),
    ProjectTelemetryPrivacyDataTag(ProjectTelemetryTag_ProductAndServicePerformance),
    TraceLoggingBoolean(TRUE, "UTCReplace_AppSessionGuid"),
    TraceLoggingKeyword(PROJECT_KEYWORD_MEASURE));
}

void Trace::EventDesktopClosed() {
  TraceLoggingWrite(
    g_hProvider,
    "MTND::Event::DesktopClosedViaLastWindowClosed",
    ProjectTelemetryPrivacyDataTag(ProjectTelemetryTag_ProductAndServicePerformance),
    TraceLoggingBoolean(TRUE, "UTCReplace_AppSessionGuid"),
    TraceLoggingKeyword(PROJECT_KEYWORD_MEASURE));
}

void Trace::ActionMaximize() {
  TraceLoggingWrite(
    g_hProvider,
    "MTND::Action::MaximizeToNewDesktop",
    ProjectTelemetryPrivacyDataTag(ProjectTelemetryTag_ProductAndServicePerformance),
    TraceLoggingBoolean(TRUE, "UTCReplace_AppSessionGuid"),
    TraceLoggingKeyword(PROJECT_KEYWORD_MEASURE));
}

void Trace::ActionRestore() {
  TraceLoggingWrite(
    g_hProvider,
    "MTND::Action::ReturnToPrimaryDesktop",
    ProjectTelemetryPrivacyDataTag(ProjectTelemetryTag_ProductAndServicePerformance),
    TraceLoggingBoolean(TRUE, "UTCReplace_AppSessionGuid"),
    TraceLoggingKeyword(PROJECT_KEYWORD_MEASURE));
}

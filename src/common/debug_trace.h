#pragma once

void init_debug_trace_runner();
void init_debug_trace_settings();
void close_debug_trace();

#ifdef _DEBUG
void _trace(const char *function, const char *file, const int line, const char *message);
void _trace_hr(const char* function, const char* file, const int line, char const* message, int value);
void _trace_last_error(const char* function, const char* file, const int line, const char* message);
void _trace_bool(const char* function, const char* file, const int line, char const* message, bool value);
void _trace_ptr(const char* function, const char* file, const int line, char const* message, void* value);

#define trace(message) _trace(__func__, __FILE__, __LINE__, (message))
#define trace_hr(value,message) _trace_hr(__func__, __FILE__, __LINE__, (message), (int) value)
#define trace_last_error(message) _trace_last_error(__func__, __FILE__, __LINE__, (message))
#define trace_bool(value,message) _trace_bool(__func__, __FILE__, __LINE__, (message), (bool) value)
#define trace_ptr(value,message) _trace_ptr(__func__, __FILE__, __LINE__, (message), (void*) value)
#else
#define trace(message)
#define trace_last_error(message)
#define trace_bool(value,message)
#define trace_ptr(value,message)
#endif

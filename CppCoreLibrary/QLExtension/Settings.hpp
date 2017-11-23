#ifndef qlex_settings_hpp
#define qlex_settings_hpp

// settings for the QLExtension project
#ifdef QLEX_EXPORTS		// inside DLL
#define QLEX_API __declspec(dllexport) 
#else		// outside DLL
#define QLEX_API __declspec(dllimport) 
#endif

#endif
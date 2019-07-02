#pragma once
#include "pch.h"

struct StreamUriResolverFromFile : winrt::implements <StreamUriResolverFromFile, winrt::Windows::Web::IUriToStreamResolver> {
  TCHAR base_path[MAX_PATH];
  winrt::Windows::Foundation::IAsyncOperation<winrt::Windows::Storage::Streams::IInputStream> UriToStreamAsync(const winrt::Windows::Foundation::Uri & uri) const;
};


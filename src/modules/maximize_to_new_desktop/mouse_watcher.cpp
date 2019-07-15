#include "pch.h"
#include "mouse_watcher.h"
#include "virtual_desktops.h"
#include <common/monitors.h>
#include <common/common.h>
#include <uiautomation.h>

namespace {
  using stdclock = std::chrono::system_clock;

  bool mouse_initialized = false;
  bool mousein_signalled = false;
  bool mousein_reset = true;
  MouseInProc mouse_in_cb;
  MouseOutProc mouse_out_cb;
  stdclock::time_point mousein_timestamp;
  stdclock::duration mousein_wait, mousein_sleep;
  HWND popup_hwnd=nullptr;
  RECT popup_rect, buttons_rect;
  RECT mouse_window_rect;
  HWND mouse_window_hwnd;

  winrt::com_ptr<IUIAutomation> ui_automation;
  winrt::com_ptr<IUIAutomationCondition> conditions_joined_top_id;
  winrt::com_ptr<IUIAutomationCondition> conditions_joined_top_name;
  winrt::com_ptr<IUIAutomationCondition> condition_true;
  
  void initialize_ui_automation() {
    conditions_joined_top_id = nullptr;
    conditions_joined_top_name = nullptr;
    condition_true = nullptr;
    ui_automation = nullptr;
    winrt::check_hresult(CoCreateInstance(CLSID_CUIAutomation,
                                          nullptr,
                                          CLSCTX_INPROC_SERVER,
                                          IID_IUIAutomation,
                                          ui_automation.put_void()));
  };

  bool mouse_in_rect(POINT mouse_pos, RECT rect) {
    return mouse_pos.x >= rect.left && mouse_pos.x <= rect.right &&
           mouse_pos.y >= rect.top  && mouse_pos.y <= rect.bottom;
  };

  // Test if mouse is over maximize button, popup window or in the trapezoid
  // between top edges of button and popup window.
  bool mouse_in_bounds(POINT mouse_pos, RECT buttons_rect, RECT popup_rect) {
    // Test if mouse is in one of the rects
    if (mouse_in_rect(mouse_pos, buttons_rect) || mouse_in_rect(mouse_pos, popup_rect))
      return true;
    // Test if mouse is in the trapezoid
    int top = buttons_rect.top;
    int bottom = popup_rect.top;
    int heigh = bottom - top;
    if (mouse_pos.y < top || mouse_pos.y > bottom || heigh == 0)
      return false;
    int top_left = buttons_rect.left;
    int top_right = buttons_rect.right;
    int bottom_left = popup_rect.left;
    int bottom_right = popup_rect.right;
    double dleft = double(bottom_left - top_left) / heigh;
    double dright = double(bottom_right - top_right) / heigh;
    int dy = mouse_pos.y - top;
    int px_left = top_left + (int)(dleft * dy);
    int px_right = top_right + (int)(dright * dy);
    return mouse_pos.x >= px_left && mouse_pos.x <= px_right;
  }

  RECT use_dwmwa_caption_strategy(HWND hwnd, std::optional<RECT>& window_rect) {
    // Try to get caption buttons by using DWMWA_CAPTION_BUTTON_BOUNDS.
    RECT result = { 0 };
    DwmGetWindowAttribute(hwnd, DWMWA_CAPTION_BUTTON_BOUNDS, &result, sizeof(RECT));

    if (result.bottom == result.top || result.left == result.right) {
      return result;
    }

    result.right += window_rect->left;
    result.left += window_rect->left;
    result.bottom += window_rect->top;
    result.top += window_rect->top;

    long width = result.right - result.left;
    result.left += width / 3;
    result.right -= width / 3;

    // If the calculated maximize button area is outside the bounds, don't accept these results.
    if (result.left > window_rect->right ||
      result.top > window_rect->bottom ||
      result.right < window_rect->left ||
      result.bottom < window_rect->top) {
      RECT zero = {0};
      return zero;
    }
    return result;
  }

  RECT use_top_right_zone_strategy(HWND hwnd, std::optional<RECT>& window_rect) {
    RECT result = { 0 };

    auto dpi = GetDpiForWindow(hwnd);
    int buttons_width = 170 * dpi / 120;
    int horizontal_padding = 20 * dpi / 120;
    int buttons_height = 50 * dpi / 120;
    result.left = window_rect->right - buttons_width + horizontal_padding;
    result.top = window_rect->top;
    result.right = window_rect->right - horizontal_padding;
    result.bottom = window_rect->top + buttons_height;
    return result;
  }

  struct UIAutomationCachedInfo {
    winrt::com_ptr<IUIAutomationElement> window_ui_element;
    double seconds_last_ui_automation_find_duration;
  };

  std::unordered_map<HWND, UIAutomationCachedInfo> custom_ui_automation_cache;

  void initialize_ui_automation_strategy() {
    initialize_ui_automation();

    if (!ui_automation)
      return;

    if (ui_automation->CreateTrueCondition(condition_true.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> condition_maximize_restore_id;
    VARIANT prop_maximize_restore_string;
    prop_maximize_restore_string.vt = VT_BSTR;
    prop_maximize_restore_string.bstrVal = SysAllocString(L"Maximize-Restore");
    if (ui_automation->CreatePropertyCondition(UIA_AutomationIdPropertyId,
                                               prop_maximize_restore_string,
                                               condition_maximize_restore_id.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> condition_maximize_restore_name;
    if (ui_automation->CreatePropertyCondition(UIA_NamePropertyId,
                                               prop_maximize_restore_string,
                                               condition_maximize_restore_name.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> condition_maximize_id;
    VARIANT prop_maximize_string;
    prop_maximize_string.vt = VT_BSTR;
    prop_maximize_string.bstrVal = SysAllocString(L"Maximize");
    if (ui_automation->CreatePropertyCondition(UIA_AutomationIdPropertyId,
                                               prop_maximize_string,
                                               condition_maximize_id.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> condition_maximize_name;
    if (ui_automation->CreatePropertyCondition(UIA_NamePropertyId,
                                               prop_maximize_string,
                                               condition_maximize_name.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> condition_restore_id;
    VARIANT prop_restore_string;
    prop_restore_string.vt = VT_BSTR;
    prop_restore_string.bstrVal = SysAllocString(L"Restore");
    if (ui_automation->CreatePropertyCondition(UIA_AutomationIdPropertyId,
                                               prop_restore_string,
                                               condition_restore_id.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> condition_restore_name;
    if (ui_automation->CreatePropertyCondition(UIA_NamePropertyId,
                                               prop_restore_string,
                                               condition_restore_name.put()) < 0 ) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> condition_restore_down_id;
    VARIANT prop_restore_down_string;
    prop_restore_down_string.vt = VT_BSTR;
    prop_restore_down_string.bstrVal = SysAllocString(L"Restore Down");
    if (ui_automation->CreatePropertyCondition(UIA_AutomationIdPropertyId,
                                               prop_restore_down_string,
                                               condition_restore_down_id.put()) < 0 ) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> condition_restore_down_name;
    if (ui_automation->CreatePropertyCondition(UIA_NamePropertyId,
                                               prop_restore_down_string,
                                               condition_restore_down_name.put()) < 0 ) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> condition_maximizerestore_id;
    VARIANT prop_maximizerestore_string;
    prop_maximizerestore_string.vt = VT_BSTR;
    prop_maximizerestore_string.bstrVal = SysAllocString(L"MaximizeRestore");
    if (ui_automation->CreatePropertyCondition(UIA_AutomationIdPropertyId,
                                               prop_maximizerestore_string,
                                               condition_maximizerestore_id.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> condition_maximizerestore_name;
    if (ui_automation->CreatePropertyCondition(UIA_NamePropertyId,
                                               prop_maximizerestore_string,
                                               condition_maximizerestore_name.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> condition_maximizerestorebutton_id;
    VARIANT prop_maximizerestorebutton_string;
    prop_maximizerestorebutton_string.vt = VT_BSTR;
    prop_maximizerestorebutton_string.bstrVal = SysAllocString(L"MaximizeRestoreButton");
    if (ui_automation->CreatePropertyCondition(UIA_AutomationIdPropertyId,
                                               prop_maximizerestorebutton_string,
                                               condition_maximizerestorebutton_id.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> condition_maximizerestorebutton_name;
    if (ui_automation->CreatePropertyCondition(UIA_NamePropertyId,
                                               prop_maximizerestorebutton_string,
                                               condition_maximizerestorebutton_name.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> conditions_joined_id_3;
    if (ui_automation->CreateOrCondition(condition_maximizerestore_id.get(),
                                         condition_maximizerestorebutton_id.get(),
                                         conditions_joined_id_3.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> conditions_joined_id_4;
    if (ui_automation->CreateOrCondition(condition_restore_down_id.get(),
                                         conditions_joined_id_3.get(),
                                         conditions_joined_id_4.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> conditions_joined_id_2;
    if (ui_automation->CreateOrCondition(condition_restore_id.get(),
                                         conditions_joined_id_4.get(),
                                         conditions_joined_id_2.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> conditions_joined_id_1;
    if (ui_automation->CreateOrCondition(condition_maximize_id.get(),
                                         conditions_joined_id_2.get(),
                                         conditions_joined_id_1.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> conditions_joined_id_top;
    if (ui_automation->CreateOrCondition(condition_maximize_restore_id.get(),
                                         conditions_joined_id_1.get(),
                                         conditions_joined_id_top.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> conditions_joined_name_3;
    if (ui_automation->CreateOrCondition(condition_maximizerestore_name.get(),
                                         condition_maximizerestorebutton_name.get(),
                                         conditions_joined_name_3.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> conditions_joined_name_4;
    if (ui_automation->CreateOrCondition(condition_restore_down_name.get(),
                                         conditions_joined_name_3.get(),
                                         conditions_joined_name_4.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> conditions_joined_name_2;
    if (ui_automation->CreateOrCondition(condition_restore_name.get(),
                                         conditions_joined_name_4.get(),
                                         conditions_joined_name_2.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> conditions_joined_name_1;
    if (ui_automation->CreateOrCondition(condition_maximize_name.get(),
                                         conditions_joined_name_2.get(),
                                         conditions_joined_name_1.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> conditions_joined_name_top;
    if (ui_automation->CreateOrCondition(condition_maximize_restore_name.get(),
                                         conditions_joined_name_1.get(),
                                         conditions_joined_name_top.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> conditions_is_offscreen;
    VARIANT prop_false_bool;
    prop_false_bool.vt = VT_BOOL;
    prop_false_bool.boolVal = VARIANT_FALSE;
    if (ui_automation->CreatePropertyCondition(UIA_IsOffscreenPropertyId,
                                               prop_false_bool,
                                               conditions_is_offscreen.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> conditions_is_enabled;
    VARIANT prop_true_bool;
    prop_true_bool.vt = VT_BOOL;
    prop_true_bool.boolVal = VARIANT_TRUE;
    if (ui_automation->CreatePropertyCondition(UIA_IsEnabledPropertyId,
                                               prop_true_bool,
                                               conditions_is_enabled.put()) < 0) {
      return;
    }
    winrt::com_ptr<IUIAutomationCondition> conditions_joined_visible_and_condition;
    if (ui_automation->CreateAndCondition(conditions_is_enabled.get(),
                                          conditions_is_offscreen.get(),
                                          conditions_joined_visible_and_condition.put()) < 0) {
      return;
    }
    if (ui_automation->CreateAndCondition(conditions_joined_id_top.get(),
                                          conditions_joined_visible_and_condition.get(),
                                          conditions_joined_top_id.put()) < 0) {
      return;
    }
    if (ui_automation->CreateAndCondition(conditions_joined_name_top.get(),
                                          conditions_joined_visible_and_condition.get(),
                                          conditions_joined_top_name.put()) < 0) {
      return;
    }
  }

  
  RECT get_bounding_rectangle_from_hwnd_UI_element(IUIAutomationElement* query_ui_element) {
    if (!query_ui_element) {
      return {};
    }
    VARIANT bounded_rect_prop;
    bounded_rect_prop.vt = VT_NULL;
    auto hr = query_ui_element->GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId, &bounded_rect_prop);
    if (FAILED(hr) || bounded_rect_prop.vt != (VT_R8 | VT_ARRAY)) {
      return {};
    }
    RECT result = { 0 };    
    DOUBLE coord_value;
    LONG pos;
    pos = 0;
    SafeArrayGetElement(bounded_rect_prop.parray, &pos, &coord_value);
    result.left = (LONG)coord_value;
    pos = 1;
    SafeArrayGetElement(bounded_rect_prop.parray, &pos, &coord_value);
    result.top = (LONG)coord_value;
    pos = 2;
    SafeArrayGetElement(bounded_rect_prop.parray, &pos, &coord_value);
    result.right = (LONG)(result.left + coord_value);
    pos = 3;
    SafeArrayGetElement(bounded_rect_prop.parray, &pos, &coord_value);
    result.bottom = (LONG)(result.top + coord_value);
    VariantClear(&bounded_rect_prop);
    return result;
  }

  winrt::com_ptr<IUIAutomationElement> find_element_ui_automation_max_depth_strategy(IUIAutomationElement* hwnd_UI_element, IUIAutomationCondition* condition, int level, int maxlevel) {
    RECT result = { 0 };
    winrt::com_ptr<IUIAutomationElement> found;
    hwnd_UI_element->FindFirst(TreeScope_Children, condition, found.put());
    if (!found && level < maxlevel) {
      winrt::com_ptr<IUIAutomationElementArray> children_array;
      auto hr = hwnd_UI_element->FindAll(TreeScope_Children, condition_true.get(), children_array.put());
      if (hr != S_OK || !children_array) {
        return found;
      };
      int children_len = 0;
      if (children_array->get_Length(&children_len) != S_OK) {
        return found;
      }
      if (level > 1 && children_len > 10) {
        // So many children. Most likely a section we won't care about.
        return found;
      }
      for (int i = 0; i < children_len; i++) {
        winrt::com_ptr<IUIAutomationElement> child;
        hr = children_array->GetElement(i, child.put());
        if (hr != S_OK || !child) {
          continue;
        }
        found = nullptr;
        found = find_element_ui_automation_max_depth_strategy(child.get(), condition, level + 1, maxlevel);
        if (found) {
          break;
        }
      }
    }
    return found;
  }

  winrt::com_ptr<IUIAutomationElement>  find_element_ui_automation_strategy(IUIAutomationElement* hwnd_UI_element, IUIAutomationCondition* condition) {
    winrt::com_ptr<IUIAutomationElement> found;
    hwnd_UI_element->FindFirst(TreeScope_Descendants, condition, found.put());
    return found;
  }

  HWND ui_automation_strategy_last_hwnd = nullptr;
  winrt::com_ptr<IUIAutomationElement> ui_automation_strategy_element_found;
  RECT use_ui_automation_strategy(HWND hwnd, std::optional<RECT>& window_rect) {
    // A strategy to find the maximize button using UIAutomation.
    /*RECT result = { 0 };
    
    HRESULT hr;
    
    */

    winrt::com_ptr<IUIAutomationElement> hwnd_UI_element;
    if (ui_automation_strategy_last_hwnd != hwnd) {
      ui_automation_strategy_last_hwnd = hwnd;
      // We haven't searched this window for the maximize button recently. Query it.
      // This is an expensive operation for some Windows and may block them.
      if (ui_automation_strategy_element_found) {
        ui_automation_strategy_element_found = nullptr;
      }
      auto hr = ui_automation->ElementFromHandle(hwnd, hwnd_UI_element.put());
      if (hr != S_OK || hwnd_UI_element == NULL) {
        return {};
      }
      UIAutomationCachedInfo cache_elem = { 0 };
      if (auto iter = custom_ui_automation_cache.find(hwnd); iter != custom_ui_automation_cache.end()) {
        // Check if the cached element for hwnd is still valid.
        cache_elem = iter->second;
        BOOL cache_valid;
        if (ui_automation->CompareElements(cache_elem.window_ui_element.get(), hwnd_UI_element.get(), &cache_valid) != S_OK) {
          return {};
        }
        if (cache_valid) {
          // Release older window element pointer and use the new instead.
          cache_elem.window_ui_element = nullptr;
          cache_elem.window_ui_element = hwnd_UI_element;
          custom_ui_automation_cache.insert_or_assign(hwnd, cache_elem);
        } else {
          cache_elem.window_ui_element = nullptr;
          custom_ui_automation_cache.erase(iter);
          cache_elem = {0};
        }
      }
      if (cache_elem.window_ui_element && cache_elem.seconds_last_ui_automation_find_duration > 0.5) {
        // Unfortunately, it takes too long to search this window. Just use another strategy.
        return {};
      }
      auto start_time = std::chrono::high_resolution_clock::now();
      ui_automation_strategy_element_found = nullptr;
      ui_automation_strategy_element_found = find_element_ui_automation_max_depth_strategy(hwnd_UI_element.get(), conditions_joined_top_id.get(), 1, 7);
      if (!ui_automation_strategy_element_found) {
        ui_automation_strategy_element_found = find_element_ui_automation_max_depth_strategy(hwnd_UI_element.get(), conditions_joined_top_name.get(), 1, 7);
      }
      auto chrono_duration = std::chrono::high_resolution_clock::now() - start_time;
      double seconds_duration = std::chrono::duration<double>(chrono_duration).count();
      if(!ui_automation_strategy_element_found) {
        //Nothing useful found. Can't keep doing this for this Window.
        cache_elem.window_ui_element = hwnd_UI_element;
        cache_elem.seconds_last_ui_automation_find_duration=seconds_duration;
        custom_ui_automation_cache.insert_or_assign(hwnd, cache_elem);
      } else {
        cache_elem.window_ui_element = hwnd_UI_element;
        cache_elem.seconds_last_ui_automation_find_duration=0;
        custom_ui_automation_cache.insert_or_assign(hwnd, cache_elem);
      }
    }
    RECT result = { 0 };
    if (ui_automation_strategy_element_found) {
      result = get_bounding_rectangle_from_hwnd_UI_element(ui_automation_strategy_element_found.get());
      if (result.left != result.right && result.bottom != result.top) {
        // Adjust top to the top of the Window. UIAutomation seems to detect the lower half of the button sometimes.
        result.top = window_rect->top;
      }
    }
    return result;
  }

  bool terminate_mouse_thread_proc = false;
  std::thread mouse_thread;
  void mouse_thread_proc() {
    initialize_ui_automation_strategy();

    while (!terminate_mouse_thread_proc) {
      std::this_thread::sleep_for(mousein_sleep);
      auto mouse_pos = get_mouse_pos();
      if (!mouse_pos) {
        if (mousein_signalled) {
          mousein_signalled = false;
          mouse_out_cb();
        }
        continue;
      }
      if (mousein_signalled) {
        auto current_mouse_window_rect = get_window_pos(mouse_window_hwnd);
        if (!IsWindowVisible(popup_hwnd) ||
            !current_mouse_window_rect ||
            *current_mouse_window_rect != mouse_window_rect ||
            !mouse_in_bounds(*mouse_pos, buttons_rect, popup_rect)) {
          mousein_signalled = false;
          mouse_out_cb();
        }
        continue;
      }
      auto mouse_window = WindowFromPoint(*mouse_pos);
      if (mouse_window == nullptr)
        continue;
      mouse_window = GetAncestor(mouse_window, GA_ROOTOWNER);
      if (GetWindowLong(mouse_window, GWL_STYLE) & WS_CHILD)
        continue;
      if (get_desktop_index_for_window(mouse_window) < 0) {
        // Couldn't get a desktop index for the window.
        // The windows most likely won't be able to move between desktops.
        continue;
      }
      auto window_rect = get_window_pos(mouse_window);
      if (!window_rect) {
        continue;
      }

      buttons_rect = use_ui_automation_strategy(mouse_window, window_rect);
      if (buttons_rect.left == buttons_rect.right || buttons_rect.bottom == buttons_rect.top) {
        buttons_rect = use_dwmwa_caption_strategy(mouse_window, window_rect);
      }
      if (buttons_rect.left == buttons_rect.right || buttons_rect.bottom == buttons_rect.top) {
        buttons_rect = use_top_right_zone_strategy(mouse_window, window_rect);
      }
      if (buttons_rect.left == buttons_rect.right || buttons_rect.bottom == buttons_rect.top) {
        continue;
      }

      if (mouse_in_rect(*mouse_pos, buttons_rect)) {
        if (mousein_reset) {
          mousein_reset = false;
          mousein_timestamp = stdclock::now();
        }
        if (stdclock::now() - mousein_timestamp > mousein_wait) {
          mousein_signalled = true;
          mouse_window_hwnd = mouse_window;
          mouse_window_rect = *window_rect;
          auto closest_monitor = get_point_monitor(*mouse_pos);
          popup_rect = mouse_in_cb(mouse_window, buttons_rect, closest_monitor.rect);
        }
      } else {
        mousein_reset = true;
      }
    }
  }
}

void start_mouse_watcher(int ms_delay, int probe_ms_delay, MouseInProc on_mouse_in, MouseOutProc on_mouse_out, HWND popup) {
  if (!mouse_initialized) {
    popup_hwnd = popup;
    mouse_initialized = true;
    mouse_in_cb = on_mouse_in;
    mouse_out_cb = on_mouse_out;
    mousein_wait = std::chrono::milliseconds(ms_delay);
    mousein_sleep = std::chrono::milliseconds(probe_ms_delay);
    terminate_mouse_thread_proc = false;
    mouse_thread = std::thread(mouse_thread_proc);
  }
}

void stop_mouse_watcher() {
  terminate_mouse_thread_proc = true;
  if (mouse_initialized) {
    mouse_thread.join();
    mouse_initialized = false;  
  }
}

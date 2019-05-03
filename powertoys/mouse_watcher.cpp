#include "pch.h"
#include "mouse_watcher.h"
#include "virtual_desktops.h"
#include "monitors.h"
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

  IUIAutomation* ui_automation = nullptr;
  IUIAutomationCondition* p_conditions_joined_top_idAutomation = NULL;
  IUIAutomationCondition* p_conditions_joined_top_nameAutomation = NULL;

  HRESULT InitializeUIAutomation(IUIAutomation **ppAutomation)
  {
    return CoCreateInstance(CLSID_CUIAutomation, NULL,
      CLSCTX_INPROC_SERVER, IID_IUIAutomation,
      reinterpret_cast<void**>(ppAutomation));
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

  RECT use_dwmwa_caption_strategy (HWND hwnd, std::optional<RECT>& window_rect) {
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

    return result;
  }

  RECT use_top_right_zone_strategy(HWND hwnd, std::optional<RECT>& window_rect) {
    RECT result = { 0 };

    auto dpi = GetDpiForWindow(hwnd);
    int buttons_width = 170 * dpi / 120;
    int horizontal_padding = 20 * dpi / 120;
    int buttons_height = 37 * dpi / 120;
    result.left = window_rect->right - buttons_width + horizontal_padding;
    result.top = window_rect->top;
    result.right = window_rect->right - horizontal_padding;
    result.bottom = window_rect->top + buttons_height;
    return result;
  }

  void initialize_ui_automation_strategy() {
    //TODO: Definitely clean this up after experimentation.
    IUIAutomationCondition* p_condition_maximize_restore_id = NULL;
    IUIAutomationCondition* p_condition_maximize_restore_name = NULL;
    VARIANT var_prop_maximize_restore_string;
    var_prop_maximize_restore_string.vt = VT_BSTR;
    var_prop_maximize_restore_string.bstrVal = SysAllocString(L"Maximize-Restore");
    IUIAutomationCondition* p_condition_maximize_id = NULL;
    IUIAutomationCondition* p_condition_maximize_name = NULL;
    VARIANT var_prop_maximize_string;
    var_prop_maximize_string.vt = VT_BSTR;
    var_prop_maximize_string.bstrVal = SysAllocString(L"Maximize");
    IUIAutomationCondition* p_condition_restore_id = NULL;
    IUIAutomationCondition* p_condition_restore_name = NULL;
    VARIANT var_prop_restore_string;
    var_prop_restore_string.vt = VT_BSTR;
    var_prop_restore_string.bstrVal = SysAllocString(L"Restore");
    IUIAutomationCondition* p_condition_maximizerestore_id = NULL;
    IUIAutomationCondition* p_condition_maximizerestore_name = NULL;
    VARIANT var_prop_maximizerestore_string;
    var_prop_maximizerestore_string.vt = VT_BSTR;
    var_prop_maximizerestore_string.bstrVal = SysAllocString(L"MaximizeRestore");
    IUIAutomationCondition* p_condition_maximizerestorebutton_id = NULL;
    IUIAutomationCondition* p_condition_maximizerestorebutton_name = NULL;
    VARIANT var_prop_maximizerestorebutton_string;
    var_prop_maximizerestorebutton_string.vt = VT_BSTR;
    var_prop_maximizerestorebutton_string.bstrVal = SysAllocString(L"MaximizeRestoreButton");
    IUIAutomationCondition* p_conditions_is_offscreen = NULL;
    VARIANT var_prop_false_bool;
    var_prop_false_bool.vt = VT_BOOL;
    var_prop_false_bool.boolVal = VARIANT_FALSE;
    IUIAutomationCondition* p_conditions_is_enabled = NULL;
    VARIANT var_prop_true_bool;
    var_prop_true_bool.vt = VT_BOOL;
    var_prop_true_bool.boolVal = VARIANT_TRUE;
    IUIAutomationCondition* p_conditions_joined_name_1 = NULL;
    IUIAutomationCondition* p_conditions_joined_name_2 = NULL;
    IUIAutomationCondition* p_conditions_joined_name_3 = NULL;
    IUIAutomationCondition* p_conditions_joined_name_top = NULL;
    IUIAutomationCondition* p_conditions_joined_id_1 = NULL;
    IUIAutomationCondition* p_conditions_joined_id_2 = NULL;
    IUIAutomationCondition* p_conditions_joined_id_3 = NULL;
    IUIAutomationCondition* p_conditions_joined_id_top = NULL;
    IUIAutomationCondition* p_conditions_joined_visible_and_condition = NULL;
    HRESULT hr;
    /*
    if (var_prop_maximize_restore_string.bstrVal == NULL) {
      goto cleanup;
    }
    */

    CoInitialize(nullptr);
    hr = InitializeUIAutomation(&ui_automation);
    if (FAILED(hr)) {
      goto cleanup;
    }

    hr = ui_automation->CreatePropertyCondition(UIA_AutomationIdPropertyId, var_prop_maximize_restore_string, &p_condition_maximize_restore_id);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreatePropertyCondition(UIA_AutomationIdPropertyId, var_prop_maximize_string, &p_condition_maximize_id);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreatePropertyCondition(UIA_AutomationIdPropertyId, var_prop_restore_string, &p_condition_restore_id);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreatePropertyCondition(UIA_AutomationIdPropertyId, var_prop_maximizerestore_string, &p_condition_maximizerestore_id);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreatePropertyCondition(UIA_AutomationIdPropertyId, var_prop_maximizerestorebutton_string, &p_condition_maximizerestorebutton_id);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreatePropertyCondition(UIA_NamePropertyId, var_prop_maximize_restore_string, &p_condition_maximize_restore_name);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreatePropertyCondition(UIA_NamePropertyId, var_prop_maximize_string, &p_condition_maximize_name);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreatePropertyCondition(UIA_NamePropertyId, var_prop_restore_string, &p_condition_restore_name);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreatePropertyCondition(UIA_NamePropertyId, var_prop_maximizerestore_string, &p_condition_maximizerestore_name);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreatePropertyCondition(UIA_NamePropertyId, var_prop_maximizerestorebutton_string, &p_condition_maximizerestorebutton_name);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreateOrCondition(p_condition_maximizerestore_id, p_condition_maximizerestorebutton_id, &p_conditions_joined_id_3);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreateOrCondition(p_condition_restore_id, p_conditions_joined_id_3, &p_conditions_joined_id_2);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreateOrCondition(p_condition_maximize_id, p_conditions_joined_id_2, &p_conditions_joined_id_1);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreateOrCondition(p_condition_maximize_restore_id, p_conditions_joined_id_1, &p_conditions_joined_id_top);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreateOrCondition(p_condition_maximizerestore_name, p_condition_maximizerestorebutton_name, &p_conditions_joined_name_3);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreateOrCondition(p_condition_restore_name, p_conditions_joined_name_3, &p_conditions_joined_name_2);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreateOrCondition(p_condition_maximize_name, p_conditions_joined_name_2, &p_conditions_joined_name_1);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreateOrCondition(p_condition_maximize_restore_name, p_conditions_joined_name_1, &p_conditions_joined_name_top);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreatePropertyCondition(UIA_IsOffscreenPropertyId, var_prop_false_bool, &p_conditions_is_offscreen);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreatePropertyCondition(UIA_IsEnabledPropertyId, var_prop_true_bool, &p_conditions_is_enabled);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreateAndCondition(p_conditions_is_enabled, p_conditions_is_offscreen, &p_conditions_joined_visible_and_condition);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreateAndCondition(p_conditions_joined_id_top, p_conditions_joined_visible_and_condition, &p_conditions_joined_top_idAutomation);
    if (FAILED(hr)) {
      goto cleanup;
    }
    hr = ui_automation->CreateAndCondition(p_conditions_joined_name_top, p_conditions_joined_visible_and_condition, &p_conditions_joined_top_nameAutomation);
    if (FAILED(hr)) {
      goto cleanup;
    }
  cleanup:
    if (p_condition_maximize_restore_id != NULL)
      p_condition_maximize_restore_id->Release();
    if (p_condition_maximize_id != NULL)
      p_condition_maximize_id->Release();
    if (p_condition_restore_id != NULL)
      p_condition_restore_id->Release();
    if (p_condition_maximizerestore_id != NULL)
      p_condition_maximizerestore_id->Release();
    if (p_condition_maximizerestorebutton_id != NULL)
      p_condition_maximizerestorebutton_id->Release();
    if (p_condition_maximize_restore_name != NULL)
      p_condition_maximize_restore_name->Release();
    if (p_condition_maximize_name != NULL)
      p_condition_maximize_name->Release();
    if (p_condition_restore_name != NULL)
      p_condition_restore_name->Release();
    if (p_condition_maximizerestore_name != NULL)
      p_condition_maximizerestore_name->Release();
    if (p_condition_maximizerestorebutton_name != NULL)
      p_condition_maximizerestorebutton_name->Release();
    if (p_conditions_joined_id_3 != NULL)
      p_conditions_joined_id_3->Release();
    if (p_conditions_joined_id_2 != NULL)
      p_conditions_joined_id_2->Release();
    if (p_conditions_joined_id_1 != NULL)
      p_conditions_joined_id_1->Release();
    if (p_conditions_joined_id_top != NULL)
      p_conditions_joined_id_top->Release();
    if (p_conditions_joined_name_3 != NULL)
      p_conditions_joined_name_3->Release();
    if (p_conditions_joined_name_2 != NULL)
      p_conditions_joined_name_2->Release();
    if (p_conditions_joined_name_1 != NULL)
      p_conditions_joined_name_1->Release();
    if (p_conditions_joined_name_top != NULL)
      p_conditions_joined_name_top->Release();
    if (p_conditions_is_offscreen != NULL)
      p_conditions_is_offscreen->Release();
    if (p_conditions_is_enabled != NULL)
      p_conditions_is_enabled->Release();
    if (p_conditions_joined_visible_and_condition != NULL)
      p_conditions_joined_visible_and_condition->Release();
    //VariantClear(&var_prop_maximize_restore_string);
    //VariantClear(&var_prop_maximize_string);
    //VariantClear(&var_prop_restore_string);
  }

  HWND ui_automation_strategy_last_hwnd = NULL;
  IUIAutomationElement* ui_automation_strategy_element_found = NULL;

  RECT get_bounding_rectangle_from_hwnd_UI_element(IUIAutomationElement* query_ui_element) {
    RECT result = { 0 };
    VARIANT varBoundedRectProp;
    varBoundedRectProp.vt = VT_NULL;
    DOUBLE coord_value;
    LONG pos;
    HRESULT hr;

    if (query_ui_element == NULL) {
      goto cleanup;
    }

    hr = query_ui_element->GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId, &varBoundedRectProp);
    if (FAILED(hr) || varBoundedRectProp.vt != (VT_R8 | VT_ARRAY)) {
      goto cleanup;
    }

    pos = 0;
    SafeArrayGetElement(varBoundedRectProp.parray, &pos, &coord_value);
    result.left = coord_value;
    pos = 1;
    SafeArrayGetElement(varBoundedRectProp.parray, &pos, &coord_value);
    result.top = coord_value;
    pos = 2;
    SafeArrayGetElement(varBoundedRectProp.parray, &pos, &coord_value);
    result.right = result.left + coord_value;
    pos = 3;
    SafeArrayGetElement(varBoundedRectProp.parray, &pos, &coord_value);
    result.bottom = result.top + coord_value;

  cleanup:
    VariantClear(&varBoundedRectProp);
    return result;
  }

  IUIAutomationElement* find_element_ui_automation_strategy(IUIAutomationElement* hwnd_UI_element, IUIAutomationCondition* pcondition) {
    RECT result = { 0 };
    IUIAutomationElement* p_found = NULL;

    hwnd_UI_element->FindFirst(TreeScope_Descendants, pcondition, &p_found);

    return p_found;
  }

  RECT use_ui_automation_strategy(HWND hwnd, std::optional<RECT>& window_rect) {
    // A strategy to find the maximize button using UIAutomation.
    RECT result = { 0 };
    IUIAutomationElement* hwnd_UI_element = NULL;
    HRESULT hr;

    if (ui_automation_strategy_last_hwnd != hwnd) {
      ui_automation_strategy_last_hwnd = hwnd;
      // We haven't searched this window for the maximize button recently. Query it.
      // This is an expensive operation for some Windows and may block them.
      if (ui_automation_strategy_element_found != NULL) {
        ui_automation_strategy_element_found->Release();
        ui_automation_strategy_element_found = NULL;
      }
      hr = ui_automation->ElementFromHandle(hwnd, &hwnd_UI_element);
      if (hr != S_OK || hwnd_UI_element == NULL) {
        goto cleanup;
      }
      ui_automation_strategy_element_found = find_element_ui_automation_strategy(hwnd_UI_element, p_conditions_joined_top_idAutomation);
      if (ui_automation_strategy_element_found == NULL) {
        ui_automation_strategy_element_found = find_element_ui_automation_strategy(hwnd_UI_element, p_conditions_joined_top_nameAutomation);
      }
    }
    if (ui_automation_strategy_element_found != NULL) {
      result = get_bounding_rectangle_from_hwnd_UI_element(ui_automation_strategy_element_found);
      if (result.left != result.right && result.bottom != result.top) {
        // Adjust top to the top of the Window. UIAutomation seems to detect the lower half of the button sometimes.
        result.top = window_rect->top;
      }
    }
  cleanup:
    if (hwnd_UI_element != NULL)
      hwnd_UI_element->Release();
    return result;
  }


  void mouse_thread_proc() {
    initialize_ui_automation_strategy();

    while (true) {
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
        if (
          !IsWindowVisible(popup_hwnd)
          || !current_mouse_window_rect
          || *current_mouse_window_rect != mouse_window_rect
          || !mouse_in_bounds(*mouse_pos, buttons_rect, popup_rect)) {
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
      if (GetCurrentDesktopGUIDIndexForWindow(mouse_window) < 0) {
        // Couldn't get a desktop index for the window.
        // The windows most likely won't be able to move between desktops.
        continue;
      }
      auto window_rect = get_window_pos(mouse_window);
      if (!window_rect) {
        continue;
      }

      buttons_rect = use_dwmwa_caption_strategy(mouse_window, window_rect);
      if (buttons_rect.left == buttons_rect.right || buttons_rect.bottom == buttons_rect.top) {
        buttons_rect = use_ui_automation_strategy(mouse_window, window_rect);
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
    std::thread(mouse_thread_proc).detach();
  }
}

// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Microsoft.PowerToys.Settings.UI.Helpers;
using Microsoft.PowerToys.Settings.UI.Library;
using Microsoft.PowerToys.Settings.UI.ViewModels;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Media;
using Windows.ApplicationModel.DataTransfer;
using WinRT;

namespace Microsoft.PowerToys.Settings.UI.Views
{
    public sealed partial class MouseWithoutBordersPage : Page
    {
        private MouseWithoutBordersViewModel ViewModel { get; set; }

        public MouseWithoutBordersPage()
        {
            var settingsUtils = new SettingsUtils();
            ViewModel = new MouseWithoutBordersViewModel(
                settingsUtils,
                SettingsRepository<GeneralSettings>.GetInstance(settingsUtils),
                SettingsRepository<MouseWithoutBordersSettings>.GetInstance(settingsUtils),
                ShellPage.SendDefaultIPCMessage);

            DataContext = ViewModel;
            InitializeComponent();
        }

        private static T GetChildOfType<T>(DependencyObject depObj, string tag)
            where T : FrameworkElement
        {
            if (depObj == null)
            {
                return null;
            }

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
            {
                var child = VisualTreeHelper.GetChild(depObj, i);

                var result = (child as T) ?? GetChildOfType<T>(child, tag);
                if (result != null && (string)result.Tag == tag)
                {
                    return result;
                }
            }

            return null;
        }

        private int GetDeviceIndex(Border b)
        {
            return b.DataContext.As<IndexedItem<string>>().Index;
        }

        private void Device_DragStarting(UIElement sender, DragStartingEventArgs args)
        {
            args.Data.RequestedOperation = DataPackageOperation.Move;
            args.Data.Properties.Add("index", GetDeviceIndex((Border)sender));
        }

        private void Device_Drop(object sender, DragEventArgs e)
        {
            e.DataView.Properties.TryGetValue("index", out object boxIndex);
            var draggedDeviceIndex = (int)boxIndex;
            var targetDeviceIndex = GetDeviceIndex((Border)e.OriginalSource);

            ViewModel.DeviceNames.Swap(draggedDeviceIndex, targetDeviceIndex);

            var itemsControl = (ItemsControl)FindName("DevicesItemsControl");
            var binding = itemsControl.GetBindingExpression(ItemsControl.ItemsSourceProperty);
            binding.UpdateSource();
        }

        private void Device_DragOver(object sender, DragEventArgs e)
        {
            e.AcceptedOperation = DataPackageOperation.Move;
        }
    }
}

using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System;
using System.ComponentModel;
using System.Threading.Tasks;

namespace BulkEditor.UI.ViewModels
{
    /// <summary>
    /// Base class for all ViewModels with common MVVM functionality
    /// </summary>
    public abstract partial class ViewModelBase : ObservableObject
    {
        [ObservableProperty]
        private bool _isBusy;

        [ObservableProperty]
        private string _busyMessage = string.Empty;

        [ObservableProperty]
        private string _title = string.Empty;

        protected ViewModelBase()
        {
            Title = GetType().Name.Replace("ViewModel", "");
        }

        /// <summary>
        /// Executes an async operation with busy state management
        /// </summary>
        protected async Task ExecuteAsync(Func<Task> operation, string? busyMessage = null)
        {
            if (IsBusy)
                return;

            IsBusy = true;
            BusyMessage = busyMessage ?? "Processing...";

            try
            {
                await operation();
            }
            finally
            {
                IsBusy = false;
                BusyMessage = string.Empty;
            }
        }

        /// <summary>
        /// Executes an async operation with result and busy state management
        /// </summary>
        protected async Task<T> ExecuteAsync<T>(Func<Task<T>> operation, string? busyMessage = null)
        {
            if (IsBusy)
                return default(T)!;

            IsBusy = true;
            BusyMessage = busyMessage ?? "Processing...";

            try
            {
                return await operation();
            }
            finally
            {
                IsBusy = false;
                BusyMessage = string.Empty;
            }
        }

        /// <summary>
        /// Called when the ViewModel is being initialized
        /// </summary>
        public virtual Task InitializeAsync()
        {
            return Task.CompletedTask;
        }

        /// <summary>
        /// Called when the ViewModel is being cleaned up
        /// </summary>
        public virtual void Cleanup()
        {
            // Override in derived classes if needed
        }
    }
}
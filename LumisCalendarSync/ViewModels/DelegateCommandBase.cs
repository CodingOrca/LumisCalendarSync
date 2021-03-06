// Copyright (c) Microsoft Corporation. All rights reserved. See License.txt in the project root for license information.

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Input;

namespace LumisCalendarSync.ViewModels
{
    /// <summary>
    /// An <see cref="ICommand"/> whose delegates can be attached for <see cref="Execute"/> and <see cref="CanExecute"/>.
    /// </summary>
    public abstract class DelegateCommandBase : ICommand
    {
        private bool myIsActive;
        private List<WeakReference> myCanExecuteChangedHandlers;

        private readonly Func<object, Task> myExecuteMethod;
        private readonly Func<object, bool> myCanExecuteMethod;

        /// <summary>
        /// Creates a new instance of a <see cref="DelegateCommandBase"/>, specifying both the execute action and the can execute function.
        /// </summary>
        /// <param name="executeMethod">The <see cref="Action"/> to execute when <see cref="ICommand.Execute"/> is invoked.</param>
        /// <param name="canExecuteMethod">The <see cref="Func{Object,Bool}"/> to invoked when <see cref="ICommand.CanExecute"/> is invoked.</param>
        protected DelegateCommandBase(Action<object> executeMethod, Func<object, bool> canExecuteMethod)
        {
            if (executeMethod == null || canExecuteMethod == null)
                throw new ArgumentNullException(nameof(executeMethod), "Delegate Command Delegates Cannot Be Null");

            myExecuteMethod = (arg) => { executeMethod(arg); return Task.Delay(0); };
            myCanExecuteMethod = canExecuteMethod;
        }

        /// <summary>
        /// Creates a new instance of a <see cref="DelegateCommandBase"/>, specifying both the Execute action as an awaitable Task and the CanExecuteFunction function.
        /// </summary>
        /// <param name="executeMethod">The <see cref="Func{Object,Task}"/> to execute when <see cref="ICommand.Execute"/> is invoked.</param>
        /// <param name="canExecuteMethod">The <see cref="Func{Object,Bool}"/> to invoked when <see cref="ICommand.CanExecute"/> is invoked.</param>
        protected DelegateCommandBase(Func<object, Task> executeMethod, Func<object, bool> canExecuteMethod)
        {
            if (executeMethod == null || canExecuteMethod == null)
                throw new ArgumentNullException(nameof(executeMethod), "Delegate Command Delegates Cannot Be Null");

            myExecuteMethod = executeMethod;
            myCanExecuteMethod = canExecuteMethod;
        }

        /// <summary>
        /// Raises <see cref="ICommand.CanExecuteChanged"/> on the UI thread so every 
        /// command invoker can re-query <see cref="ICommand.CanExecute"/>.
        /// </summary>
        private void OnCanExecuteChanged()
        {
            WeakEventHandlerManager.CallWeakReferenceHandlers(this, myCanExecuteChangedHandlers);
        }


        /// <summary>
        /// Raises <see cref="DelegateCommandBase.CanExecuteChanged"/> on the UI thread so every command invoker
        /// can re-query to check if the command can execute.
        /// <remarks>Note that this will trigger the execution of <see cref="DelegateCommandBase.CanExecute"/> once for each invoker.</remarks>
        /// </summary>
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1030:UseEventsWhereAppropriate")]
        public void RaiseCanExecuteChanged()
        {
            OnCanExecuteChanged();
        }

        async void ICommand.Execute(object parameter)
        {
            await Execute(parameter);
        }

        bool ICommand.CanExecute(object parameter)
        {
            return CanExecute(parameter);
        }

        /// <summary>
        /// Executes the command with the provided parameter by invoking the <see cref="Action{Object}"/> supplied during construction.
        /// </summary>
        /// <param name="parameter"></param>
        protected async Task Execute(object parameter)
        {
            await myExecuteMethod(parameter);
        }

        /// <summary>
        /// Determines if the command can execute with the provided parameter by invoking the <see cref="Func{Object,Bool}"/> supplied during construction.
        /// </summary>
        /// <param name="parameter">The parameter to use when determining if this command can execute.</param>
        /// <returns>Returns <see langword="true"/> if the command can execute.  <see langword="False"/> otherwise.</returns>
        protected bool CanExecute(object parameter)
        {
            return myCanExecuteMethod == null || myCanExecuteMethod(parameter);
        }

        /// <inheritdoc />
        /// <summary>
        /// Occurs when changes occur that affect whether or not the command should execute. You must keep a hard
        /// reference to the handler to avoid garbage collection and unexpected results. See remarks for more information.
        /// </summary>
        /// <remarks>
        /// When subscribing to the <see cref="E:System.Windows.Input.ICommand.CanExecuteChanged" /> event using 
        /// code (not when binding using XAML) will need to keep a hard reference to the event handler. This is to prevent 
        /// garbage collection of the event handler because the command implements the Weak Event pattern so it does not have
        /// a hard reference to this handler. An example implementation can be seen in the CompositeCommand and CommandBehaviorBase
        /// classes. In most scenarios, there is no reason to sign up to the CanExecuteChanged event directly, but if you do, you
        /// are responsible for maintaining the reference.
        /// </remarks>
        /// <example>
        /// The following code holds a reference to the event handler. The myEventHandlerReference value should be stored
        /// in an instance member to avoid it from being garbage collected.
        /// <code>
        /// EventHandler myEventHandlerReference = new EventHandler(this.OnCanExecuteChanged);
        /// command.CanExecuteChanged += myEventHandlerReference;
        /// </code>
        /// </example>
        public virtual event EventHandler CanExecuteChanged
        {
            add => WeakEventHandlerManager.AddWeakReferenceHandler(ref myCanExecuteChangedHandlers, value, 2);
            remove => WeakEventHandlerManager.RemoveWeakReferenceHandler(myCanExecuteChangedHandlers, value);
        }

        #region IsActive
        /// <summary>
        /// Gets or sets a value indicating whether the object is active.
        /// </summary>
        /// <value><see langword="true" /> if the object is active; otherwise <see langword="false" />.</value>
        public bool IsActive
        {
            get => myIsActive;
            set
            {
                if (myIsActive == value) return;
                myIsActive = value;
                OnIsActiveChanged();
            }
        }

        /// <summary>
        /// Fired if the <see cref="IsActive"/> property changes.
        /// </summary>
        public event EventHandler IsActiveChanged;

        /// <summary>
        /// This raises the <see cref="DelegateCommandBase.IsActiveChanged"/> event.
        /// </summary>
        private void OnIsActiveChanged()
        {
            IsActiveChanged?.Invoke(this, EventArgs.Empty);
        }
        #endregion
    }
}
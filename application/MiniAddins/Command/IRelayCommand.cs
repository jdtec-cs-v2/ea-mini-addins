using System;
using System.Windows.Input;

namespace MiniAddins.Command
{
    public interface IRelayCommand:ICommand
    {
        event EventHandler Executed;
        event EventHandler Executing;
    }
}

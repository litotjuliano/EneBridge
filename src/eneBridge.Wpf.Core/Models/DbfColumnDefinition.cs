namespace eneBridge.Wpf.Core.Models;

public sealed record DbfColumnDefinition(string Name, Type ClrType, int MaxLength = 0);

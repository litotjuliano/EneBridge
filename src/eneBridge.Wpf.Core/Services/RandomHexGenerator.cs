using System.Security.Cryptography;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Generates the 8-character uppercase hex correlation id used for ictrane's entry2 column.
/// The original used an unseeded System.Random; this uses a cryptographic RNG instead for a
/// cleaner guaranteed-uppercase-hex result — functionally equivalent (still just a correlation
/// id, not security-sensitive), same 8-hex-char output shape.
/// </summary>
public static class RandomHexGenerator
{
    public static string Generate() => Convert.ToHexString(RandomNumberGenerator.GetBytes(4));
}

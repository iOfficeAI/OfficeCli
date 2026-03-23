// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Reflection;

namespace OfficeCli.Core;

/// <summary>
/// Installs officecli skills (SKILL.md) into AI client skill directories.
/// </summary>
public static class SkillInstaller
{
    private static readonly Dictionary<string, string> SkillResources = new()
    {
        ["officecli"] = "OfficeCli.Resources.skill-officecli.md",
    };

    public static void Install(string target)
    {
        switch (target.ToLowerInvariant())
        {
            case "claude" or "claude-code":
                InstallClaude();
                break;
            default:
                Console.Error.WriteLine($"Unknown target: {target}");
                Console.Error.WriteLine("Supported: claude (Claude Code)");
                break;
        }
    }

    private static void InstallClaude()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var skillsDir = Path.Combine(home, ".claude", "skills");
        var installed = new List<string>();

        foreach (var (skillName, resourceName) in SkillResources)
        {
            var content = LoadEmbeddedResource(resourceName);
            if (content == null)
            {
                Console.Error.WriteLine($"  Warning: embedded resource not found for skill '{skillName}'");
                continue;
            }

            var targetDir = Path.Combine(skillsDir, skillName);
            Directory.CreateDirectory(targetDir);
            var targetPath = Path.Combine(targetDir, "SKILL.md");

            // Check if already up to date
            if (File.Exists(targetPath) && File.ReadAllText(targetPath) == content)
            {
                Console.WriteLine($"  {skillName}: already up to date");
                installed.Add(skillName);
                continue;
            }

            File.WriteAllText(targetPath, content);
            Console.WriteLine($"  {skillName}: installed");
            installed.Add(skillName);
        }

        if (installed.Count > 0)
        {
            Console.WriteLine();
            Console.WriteLine($"Skills installed to {skillsDir}");
            Console.WriteLine("These skills are now available in all Claude Code projects.");
        }
    }

    private static string? LoadEmbeddedResource(string resourceName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        using var stream = assembly.GetManifestResourceStream(resourceName);
        if (stream == null) return null;
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}

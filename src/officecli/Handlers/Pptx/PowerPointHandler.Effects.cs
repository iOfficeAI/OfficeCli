// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    /// <summary>
    /// Apply outer shadow effect to ShapeProperties.
    /// Format: "COLOR" or "COLOR-BLUR-ANGLE-DIST" or "COLOR-BLUR-ANGLE-DIST-OPACITY"
    ///   COLOR: hex (e.g. 000000)
    ///   BLUR: blur radius in points, default 4
    ///   ANGLE: direction in degrees, default 45
    ///   DIST: distance in points, default 3
    ///   OPACITY: 0-100 percent, default 40
    /// Examples: "000000", "000000-6-315-4-50", "none"
    /// </summary>
    private static void ApplyShadow(ShapeProperties spPr, string value)
    {
        var effectList = spPr.GetFirstChild<Drawing.EffectList>() ?? spPr.AppendChild(new Drawing.EffectList());
        effectList.RemoveAllChildren<Drawing.OuterShadow>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) || value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (!effectList.HasChildren) spPr.RemoveChild(effectList);
            return;
        }

        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("Shadow value cannot be empty. Use 'none' to remove shadow.");

        var parts = value.Split('-');
        var colorHex = parts[0].TrimStart('#').ToUpperInvariant();
        if (!double.TryParse(parts.Length > 1 ? parts[1] : "4", out var blurPt))
            throw new ArgumentException($"Invalid shadow blur value: '{parts[1]}'. Expected a number.");
        if (!double.TryParse(parts.Length > 2 ? parts[2] : "45", out var angleDeg))
            throw new ArgumentException($"Invalid shadow angle value: '{parts[2]}'. Expected a number.");
        if (!double.TryParse(parts.Length > 3 ? parts[3] : "3", out var distPt))
            throw new ArgumentException($"Invalid shadow distance value: '{parts[3]}'. Expected a number.");
        if (!double.TryParse(parts.Length > 4 ? parts[4] : "40", out var opacity))
            throw new ArgumentException($"Invalid shadow opacity value: '{parts[4]}'. Expected a number.");

        var shadow = new Drawing.OuterShadow
        {
            BlurRadius    = (long)(blurPt * 12700),
            Distance      = (long)(distPt * 12700),
            Direction     = (int)(angleDeg * 60000),
            Alignment     = Drawing.RectangleAlignmentValues.TopLeft,
            RotateWithShape = false
        };
        var clr = BuildColorElement(parts[0]);
        clr.AppendChild(new Drawing.Alpha { Val = (int)(opacity * 1000) });
        shadow.AppendChild(clr);
        effectList.AppendChild(shadow);
    }

    /// <summary>
    /// Apply glow effect to ShapeProperties.
    /// Format: "COLOR" or "COLOR-RADIUS" or "COLOR-RADIUS-OPACITY"
    ///   COLOR: hex (e.g. 0070FF)
    ///   RADIUS: glow radius in points, default 8
    ///   OPACITY: 0-100 percent, default 75
    /// Examples: "0070FF", "FF0000-10", "00B0F0-6-60", "none"
    /// </summary>
    private static void ApplyGlow(ShapeProperties spPr, string value)
    {
        var effectList = spPr.GetFirstChild<Drawing.EffectList>() ?? spPr.AppendChild(new Drawing.EffectList());
        effectList.RemoveAllChildren<Drawing.Glow>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) || value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (!effectList.HasChildren) spPr.RemoveChild(effectList);
            return;
        }

        var parts = value.Split('-');
        var colorHex = parts[0].TrimStart('#').ToUpperInvariant();
        if (!double.TryParse(parts.Length > 1 ? parts[1] : "8", out var radiusPt))
            throw new ArgumentException($"Invalid glow radius value: '{parts[1]}'. Expected a number.");
        if (!double.TryParse(parts.Length > 2 ? parts[2] : "75", out var opacity))
            throw new ArgumentException($"Invalid glow opacity value: '{parts[2]}'. Expected a number.");

        var glow = new Drawing.Glow { Radius = (long)(radiusPt * 12700) };
        var glowClr = BuildColorElement(parts[0]);
        glowClr.AppendChild(new Drawing.Alpha { Val = (int)(opacity * 1000) });
        glow.AppendChild(glowClr);
        effectList.AppendChild(glow);
    }

    /// <summary>
    /// Apply reflection effect to ShapeProperties.
    /// Format: "TYPE" where TYPE is one of:
    ///   tight / small  — tight reflection, touching (stA=52000 endA=300 endPos=55000)
    ///   half           — half reflection (stA=52000 endA=300 endPos=90000)
    ///   full           — full reflection (stA=52000 endA=300 endPos=100000)
    ///   true           — alias for half
    ///   none / false   — remove reflection
    /// </summary>
    private static void ApplyReflection(ShapeProperties spPr, string value)
    {
        var effectList = spPr.GetFirstChild<Drawing.EffectList>() ?? spPr.AppendChild(new Drawing.EffectList());
        effectList.RemoveAllChildren<Drawing.Reflection>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) || value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (!effectList.HasChildren) spPr.RemoveChild(effectList);
            return;
        }

        // endPos controls how much of the shape is reflected
        int endPos = value.ToLowerInvariant() switch
        {
            "tight" or "small" => 55000,
            "true" or "half"   => 90000,
            "full"             => 100000,
            _ => int.TryParse(value, out var pct) ? (int)Math.Min((long)pct * 1000, 100000) : 90000
        };

        var reflection = new Drawing.Reflection
        {
            BlurRadius      = 6350,
            StartOpacity    = 52000,
            StartPosition   = 0,
            EndAlpha        = 300,
            EndPosition     = endPos,
            Distance        = 0,
            Direction       = 5400000,  // 90° — downward
            VerticalRatio   = -100000,  // flip vertically
            Alignment       = Drawing.RectangleAlignmentValues.BottomLeft,
            RotateWithShape = false
        };
        effectList.AppendChild(reflection);
    }

    /// <summary>
    /// Apply soft edge effect to ShapeProperties.
    /// Value: radius in points (e.g. "5") or "none" to remove.
    /// </summary>
    private static void ApplySoftEdge(ShapeProperties spPr, string value)
    {
        var effectList = spPr.GetFirstChild<Drawing.EffectList>() ?? spPr.AppendChild(new Drawing.EffectList());
        effectList.RemoveAllChildren<Drawing.SoftEdge>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) || value.Equals("false", StringComparison.OrdinalIgnoreCase))
        {
            if (!effectList.HasChildren) spPr.RemoveChild(effectList);
            return;
        }

        var radiusPt = double.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
        effectList.AppendChild(new Drawing.SoftEdge { Radius = (long)(radiusPt * 12700) });
    }

    /// <summary>
    /// Apply 3D rotation (scene3d) to ShapeProperties.
    /// Format: "rotX,rotY,rotZ" in degrees (e.g. "45,30,0")
    /// </summary>
    private static void Apply3DRotation(ShapeProperties spPr, string value)
    {
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            var existing = spPr.GetFirstChild<Drawing.Scene3DType>();
            if (existing != null) spPr.RemoveChild(existing);
            return;
        }

        var parts = value.Split(',');
        var rotX = double.Parse(parts[0].Trim(), System.Globalization.CultureInfo.InvariantCulture);
        var rotY = parts.Length > 1 ? double.Parse(parts[1].Trim(), System.Globalization.CultureInfo.InvariantCulture) : 0;
        var rotZ = parts.Length > 2 ? double.Parse(parts[2].Trim(), System.Globalization.CultureInfo.InvariantCulture) : 0;

        var scene3d = EnsureScene3D(spPr);
        var camera = scene3d.Camera!;
        camera.Rotation = new Drawing.Rotation
        {
            Latitude = (int)(rotX * 60000),
            Longitude = (int)(rotY * 60000),
            Revolution = (int)(rotZ * 60000)
        };
    }

    /// <summary>
    /// Apply a single 3D rotation axis.
    /// </summary>
    private static void Apply3DRotationAxis(ShapeProperties spPr, string axis, string value)
    {
        var scene3d = EnsureScene3D(spPr);
        var camera = scene3d.Camera!;
        var rot = camera.Rotation ?? (camera.Rotation = new Drawing.Rotation());
        var deg = (int)(double.Parse(value, System.Globalization.CultureInfo.InvariantCulture) * 60000);

        switch (axis)
        {
            case "x": rot.Latitude = deg; break;
            case "y": rot.Longitude = deg; break;
            case "z": rot.Revolution = deg; break;
        }
    }

    /// <summary>
    /// Apply bevel to ShapeProperties (top or bottom).
    /// Format: "preset" or "preset-width-height" (width/height in points)
    /// Presets: circle, relaxedInset, cross, coolSlant, angle, softRound, convex,
    ///          slope, divot, riblet, hardEdge, artDeco
    /// Examples: "circle", "circle-6-6", "none"
    /// </summary>
    private static void ApplyBevel(ShapeProperties spPr, string value, bool top)
    {
        var sp3d = spPr.GetFirstChild<Drawing.Shape3DType>();

        if (value.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            if (sp3d != null)
            {
                if (top) { sp3d.BevelTop = null; }
                else { sp3d.BevelBottom = null; }
                if (sp3d.BevelTop == null && sp3d.BevelBottom == null &&
                    (sp3d.ExtrusionHeight == null || sp3d.ExtrusionHeight.Value == 0))
                    spPr.RemoveChild(sp3d);
            }
            return;
        }

        sp3d ??= EnsureShape3D(spPr);
        var bevelParts = value.Split('-');
        var preset = ParseBevelPreset(bevelParts[0].Trim());
        var w = bevelParts.Length > 1 ? (long)(double.Parse(bevelParts[1].Trim(), System.Globalization.CultureInfo.InvariantCulture) * 12700) : 76200L; // default 6pt
        var h = bevelParts.Length > 2 ? (long)(double.Parse(bevelParts[2].Trim(), System.Globalization.CultureInfo.InvariantCulture) * 12700) : w;

        if (top)
        {
            sp3d.BevelTop = new Drawing.BevelTop { Width = w, Height = h, Preset = preset };
        }
        else
        {
            sp3d.BevelBottom = new Drawing.BevelBottom { Width = w, Height = h, Preset = preset };
        }
    }

    /// <summary>
    /// Apply 3D extrusion depth in points.
    /// </summary>
    private static void Apply3DDepth(ShapeProperties spPr, string value)
    {
        if (value.Equals("none", StringComparison.OrdinalIgnoreCase) || value == "0")
        {
            var sp3d = spPr.GetFirstChild<Drawing.Shape3DType>();
            if (sp3d != null) { sp3d.ExtrusionHeight = 0; }
            return;
        }

        var sp3dEl = EnsureShape3D(spPr);
        sp3dEl.ExtrusionHeight = (long)(double.Parse(value, System.Globalization.CultureInfo.InvariantCulture) * 12700);
    }

    /// <summary>
    /// Apply 3D material preset.
    /// </summary>
    private static void Apply3DMaterial(ShapeProperties spPr, string value)
    {
        var sp3d = EnsureShape3D(spPr);
        sp3d.PresetMaterial = ParseMaterial(value);
    }

    /// <summary>
    /// Apply light rig preset to scene3d.
    /// </summary>
    private static void ApplyLightRig(ShapeProperties spPr, string value)
    {
        var scene3d = EnsureScene3D(spPr);
        scene3d.LightRig!.Rig = ParseLightRig(value);
    }

    // --- Helper methods ---

    private static Drawing.Scene3DType EnsureScene3D(ShapeProperties spPr)
    {
        var scene3d = spPr.GetFirstChild<Drawing.Scene3DType>();
        if (scene3d != null) return scene3d;

        scene3d = new Drawing.Scene3DType(
            new Drawing.Camera { Preset = Drawing.PresetCameraValues.OrthographicFront },
            new Drawing.LightRig { Rig = Drawing.LightRigValues.ThreePoints, Direction = Drawing.LightRigDirectionValues.Top }
        );
        spPr.AppendChild(scene3d);
        return scene3d;
    }

    private static Drawing.Shape3DType EnsureShape3D(ShapeProperties spPr)
    {
        var sp3d = spPr.GetFirstChild<Drawing.Shape3DType>();
        if (sp3d != null) return sp3d;

        sp3d = new Drawing.Shape3DType();
        spPr.AppendChild(sp3d);
        return sp3d;
    }

    private static Drawing.BevelPresetValues ParseBevelPreset(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "circle" => Drawing.BevelPresetValues.Circle,
            "relaxedinset" => Drawing.BevelPresetValues.RelaxedInset,
            "cross" => Drawing.BevelPresetValues.Cross,
            "coolslant" => Drawing.BevelPresetValues.CoolSlant,
            "angle" => Drawing.BevelPresetValues.Angle,
            "softround" => Drawing.BevelPresetValues.SoftRound,
            "convex" => Drawing.BevelPresetValues.Convex,
            "slope" => Drawing.BevelPresetValues.Slope,
            "divot" => Drawing.BevelPresetValues.Divot,
            "riblet" => Drawing.BevelPresetValues.Riblet,
            "hardedge" => Drawing.BevelPresetValues.HardEdge,
            "artdeco" => Drawing.BevelPresetValues.ArtDeco,
            _ => Drawing.BevelPresetValues.Circle
        };
    }

    private static Drawing.PresetMaterialTypeValues ParseMaterial(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "warmmatte" => Drawing.PresetMaterialTypeValues.WarmMatte,
            "plastic" => Drawing.PresetMaterialTypeValues.Plastic,
            "metal" => Drawing.PresetMaterialTypeValues.Metal,
            "dkedge" or "darkedge" => Drawing.PresetMaterialTypeValues.DarkEdge,
            "softedge" => Drawing.PresetMaterialTypeValues.SoftEdge,
            "flat" => Drawing.PresetMaterialTypeValues.Flat,
            "wire" or "wireframe" => Drawing.PresetMaterialTypeValues.LegacyWireframe,
            "powder" => Drawing.PresetMaterialTypeValues.Powder,
            "translucentpowder" => Drawing.PresetMaterialTypeValues.TranslucentPowder,
            "clear" => Drawing.PresetMaterialTypeValues.Clear,
            "softmetal" => Drawing.PresetMaterialTypeValues.SoftMetal,
            "matte" => Drawing.PresetMaterialTypeValues.Matte,
            _ => Drawing.PresetMaterialTypeValues.Plastic
        };
    }

    private static Drawing.LightRigValues ParseLightRig(string value)
    {
        return value.ToLowerInvariant() switch
        {
            "threept" or "3pt" => Drawing.LightRigValues.ThreePoints,
            "balanced" => Drawing.LightRigValues.Balanced,
            "soft" => Drawing.LightRigValues.Soft,
            "harsh" => Drawing.LightRigValues.Harsh,
            "flood" => Drawing.LightRigValues.Flood,
            "contrasting" => Drawing.LightRigValues.Contrasting,
            "morning" => Drawing.LightRigValues.Morning,
            "sunrise" => Drawing.LightRigValues.Sunrise,
            "sunset" => Drawing.LightRigValues.Sunset,
            "chilly" => Drawing.LightRigValues.Chilly,
            "freezing" => Drawing.LightRigValues.Freezing,
            "flat" => Drawing.LightRigValues.Flat,
            "twopt" or "2pt" => Drawing.LightRigValues.TwoPoints,
            "glow" => Drawing.LightRigValues.Glow,
            "brightroom" => Drawing.LightRigValues.BrightRoom,
            _ => Drawing.LightRigValues.ThreePoints
        };
    }

    /// <summary>
    /// Format a bevel element as "preset-width-height" string for reading back.
    /// </summary>
    internal static string FormatBevel(Drawing.BevelType bevel)
    {
        var preset = bevel.Preset?.HasValue == true ? bevel.Preset.InnerText : "circle";
        var w = bevel.Width?.HasValue == true ? $"{bevel.Width.Value / 12700.0:0.##}" : "6";
        var h = bevel.Height?.HasValue == true ? $"{bevel.Height.Value / 12700.0:0.##}" : "6";
        return $"{preset}-{w}-{h}";
    }
}

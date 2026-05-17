// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // PR1 stub. PR2 fills in: walk the slide's NotesSlidePart, emit
    // `add notes parent="/slide[N]"` + the notes body's paragraphs/runs
    // through the same EmitTextBody pipeline used for slide shapes.
    private static void EmitNotes(PowerPointHandler ppt, string slidePath,
                                  List<BatchItem> items, SlideEmitContext ctx)
    {
        _ = ppt; _ = slidePath; _ = items; _ = ctx;
    }
}

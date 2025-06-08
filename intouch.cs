
 bool heightOk = Math.Abs(actualHeight - A4HeightTwips) <= PageSizeTolerance;

    if (!widthOk || !heightOk)
    {
        string widthStr = $"{actualWidth} twipsów (~{actualWidth / 56.6929:0.##} mm)";
        string heightStr = $"{actualHeight} twipsów (~{actualHeight / 56.6929:0.##} mm)";

        issues.Add(new(WcagIssue(WcagErrorCode.InvalidPageSize,
            $"Rozmiar strony różni się od A4. Oczekiwano 210x297 mm, wykryto {widthStr} x {heightStr}.")));
    }

    return issues;

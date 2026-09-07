using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace ChuanHoa.Client.Core.Tests;

public sealed class AdministrativeReferenceDataTests
{
    [Fact]
    public void Current_administrative_snapshot_matches_the_2025_two_level_model()
    {
        using var document = JsonDocument.Parse(File.ReadAllText(DictionaryPath("administrative_units.json")));
        var active = document.RootElement.EnumerateArray()
            .Where(item => item.GetProperty("status").GetString() == "Active")
            .ToArray();
        var provinces = active.Where(item => item.GetProperty("level").GetString() == "Province").ToArray();
        var communes = active.Where(item => item.GetProperty("level").GetString() == "Commune").ToArray();

        Assert.Equal(34, provinces.Length);
        Assert.Equal(3321, communes.Length);
        Assert.DoesNotContain(active, item => item.GetProperty("level").GetString() == "District");
        Assert.Equal(34, provinces.Select(item => item.GetProperty("code").GetString()).Distinct().Count());
        Assert.Equal(3321, communes.Select(item => item.GetProperty("code").GetString()).Distinct().Count());

        var names = provinces.ToDictionary(item => item.GetProperty("code").GetString()!,
            item => item.GetProperty("canonicalName").GetString()!);
        Assert.Equal("Thành phố Hà Nội", names["01"]);
        Assert.Equal("Thành phố Huế", names["46"]);
        Assert.Equal("Thành phố Đà Nẵng", names["48"]);
        Assert.Equal("Thành phố Hồ Chí Minh", names["79"]);
        Assert.Equal("Thành phố Cần Thơ", names["92"]);

        var provinceCodes = provinces.Select(item => item.GetProperty("code").GetString()).ToHashSet();
        Assert.All(communes, item => Assert.Contains(item.GetProperty("provinceCode").GetString(), provinceCodes));
    }

    [Fact]
    public void Historical_province_aliases_are_not_active_units()
    {
        using var historical = JsonDocument.Parse(File.ReadAllText(
            DictionaryPath("historical_administrative_aliases.json")));
        Assert.NotEmpty(historical.RootElement.EnumerateArray());
        Assert.All(historical.RootElement.EnumerateArray(), item =>
            Assert.Equal("Historical", item.GetProperty("status").GetString()));
        Assert.Contains(historical.RootElement.EnumerateArray(), item =>
            item.GetProperty("name").GetString() == "Tỉnh Hà Giang" &&
            item.GetProperty("successorCode").GetString() == "08");
    }

    private static string DictionaryPath(string fileName)
    {
        var copied = Path.Combine(AppContext.BaseDirectory, "ReferenceData", fileName);
        if (File.Exists(copied)) return copied;
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null)
        {
            var candidate = Path.Combine(directory.FullName, "shared", "dictionaries", fileName);
            if (File.Exists(candidate)) return candidate;
            directory = directory.Parent;
        }
        var workingCandidate = Path.Combine(Directory.GetCurrentDirectory(), "shared", "dictionaries", fileName);
        if (File.Exists(workingCandidate)) return workingCandidate;
        throw new FileNotFoundException("Repository dictionary was not found.", workingCandidate);
    }
}

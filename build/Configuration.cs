using System;
using System.ComponentModel;
using System.Linq;
using Nuke.Common.Tooling;

[TypeConverter(typeof(TypeConverter<Configuration>))]
public class Configuration : Enumeration
{
    public static Configuration Release_2018 = new Configuration { Value = nameof(Release_2018) };
    public static Configuration Release_2019 = new Configuration { Value = nameof(Release_2019) };
    public static Configuration Release_2020 = new Configuration { Value = nameof(Release_2020) };
    public static Configuration Release_2021 = new Configuration { Value = nameof(Release_2021) };
    public static Configuration Release_2022 = new Configuration { Value = nameof(Release_2022) };
    public static Configuration Release_2023 = new Configuration { Value = nameof(Release_2023) };
    public static Configuration Release_2024 = new Configuration { Value = nameof(Release_2024) };
    public static Configuration Release_2025 = new Configuration { Value = nameof(Release_2025) };

    public static Configuration Release = new Configuration { Value = nameof(Release) };

    public static implicit operator string(Configuration configuration) => configuration.Value;
}

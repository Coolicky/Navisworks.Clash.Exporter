using Nuke.Common;
using Nuke.Common.ProjectModel;
using Nuke.Common.Tools.MSBuild;

class Build : NukeBuild
{
    public static int Main () => Execute<Build>(x => x.Release);

    [Parameter] readonly Configuration Configuration = Configuration.Release;
    [Solution] readonly Solution Solution;

    Target Clean => _ => _
        .Executes(() =>
        {
        });

    Target Release => _ => _
        .DependsOn(Clean)
        .Executes(() =>
        {
            BuildPlugin();
            BuildConfig(Configuration.Release);
        });

    void BuildPlugin()
    {
        BuildConfig(Configuration.Release_2018);
        BuildConfig(Configuration.Release_2019);
        BuildConfig(Configuration.Release_2020);
        BuildConfig(Configuration.Release_2021);
        BuildConfig(Configuration.Release_2022);
        BuildConfig(Configuration.Release_2023);
        BuildConfig(Configuration.Release_2024);
        BuildConfig(Configuration.Release_2025);
    }

    void BuildConfig(Configuration config) =>
        MSBuildTasks.MSBuild(r => r
            .EnableRestore()
            .SetTargetPath(Solution)
            .SetConfiguration(config));
}

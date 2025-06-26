using Helpers;
using Models;

public class FixturesTests
{
    [Fact]
    public void AllTeamsPlayAtLeastOnce()
    {
        // Arrange: Example teams and dummy players
        var teamNames = new List<string> { "A", "B", "C", "D" };
        var teamsDict = teamNames.ToDictionary(
            t => t,
            t => new List<ExcelPlayerModel> { new ExcelPlayerModel
                {
                    Name = t,
                    IsPlayer = true,
                    Team = t,
                    IsSub = false,
                    Division = 0,
                    Handicap = 0
                }
            }
        );

        // Act: Generate fixtures
        var fixtures = Fixtures.GenerateFixtures(teamNames, teamsDict, randomize: false);

        // Collect all teams that played (exclude BYE)
        var playedTeams = new HashSet<string>();
        foreach (var fixture in fixtures)
        {
            foreach (var team in fixture.Teams.Keys)
            {
                if (!team.EndsWith("(BYE)"))
                    playedTeams.Add(team);
            }
        }

        // Assert: Every team has played at least once
        foreach (var team in teamNames)
        {
            Assert.Contains(team, playedTeams);
        }
    }
}
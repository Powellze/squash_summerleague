using Models;

namespace Helpers
{
    public static class Fixtures
    {
        public static List<FixtureModel<ExcelPlayerModel>> GenerateFixtures(List<string> teamNames, Dictionary<string, List<ExcelPlayerModel>> teamsDict, bool randomize = false)
        {
            int numTeams = teamNames.Count;
            bool isOdd = numTeams % 2 != 0;
            var teams = new List<string>(teamNames);

            if (randomize)
            {
                var rng = new Random();
                teams = teams.OrderBy(_ => rng.Next()).ToList();
            }

            if (isOdd)
                teams.Add("BYE");

            int totalRounds = teams.Count - 1;
            int matchesPerRound = teams.Count / 2;
            var fixtures = new List<FixtureModel<ExcelPlayerModel>>();

            for (int round = 0; round < totalRounds; round++)
            {
                var weekTeams = new Dictionary<string, List<ExcelPlayerModel>>();

                for (int match = 0; match < matchesPerRound; match++)
                {
                    int home = (round + match) % (teams.Count - 1);
                    int away = (teams.Count - 1 - match + round) % (teams.Count - 1);

                    if (match == 0)
                        away = teams.Count - 1;

                    string teamA = teams[home];
                    string teamB = teams[away];

                    // Only add a match if neither team is BYE
                    if (teamA != "BYE" && teamB != "BYE")
                    {
                        weekTeams[teamA] = teamsDict[teamA];
                        weekTeams[teamB] = teamsDict[teamB];
                    }
                    // If one team is BYE, only add the other team as having a bye
                    else if (teamA == "BYE" && teamB != "BYE")
                    {
                        weekTeams[teamB + " (BYE)"] = teamsDict[teamB];
                    }
                    else if (teamB == "BYE" && teamA != "BYE")
                    {
                        weekTeams[teamA + " (BYE)"] = teamsDict[teamA];
                    }
                    // If both are BYE, do nothing
                }

                fixtures.Add(new FixtureModel<ExcelPlayerModel>
                {
                    Week = round + 1,
                    Teams = weekTeams
                });
            }

            return fixtures;
        }
    }
}

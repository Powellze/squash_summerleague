namespace Models;

public class FixtureModel<TPlayer>
{
    public int Week { get; set; }
    public Dictionary<string, List<TPlayer>> Teams { get; set; } = new();
}
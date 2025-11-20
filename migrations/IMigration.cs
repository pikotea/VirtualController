public interface IMigration
{
    string Id { get; }
    string Description { get; }
    void Up();
}
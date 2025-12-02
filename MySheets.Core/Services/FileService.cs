using System.Text.Json;
using MySheets.Core.Models;

namespace MySheets.Core.Services;

public record CellDto(int Row, int Col, string Expression);

public class FileService {
    public void Save(string path, IEnumerable<Cell> cells) {
        var data = cells.Select(c => new CellDto(c.Row, c.Col, c.Expression));
        var options = new JsonSerializerOptions { WriteIndented = true };
        var json = JsonSerializer.Serialize(data, options);
        File.WriteAllText(path, json);
    }

    public IEnumerable<CellDto> Load(string path) {
        if (!File.Exists(path)) return new List<CellDto>();
        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<IEnumerable<CellDto>>(json) ?? new List<CellDto>();
    }
}
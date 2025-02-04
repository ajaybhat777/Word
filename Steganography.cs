using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

public class AlphaSteganography
{
    public string ExtractHiddenData(byte[] imageBytes)
    {
        using var image = Image.Load<Rgba32>(imageBytes);
        var bits = new List<char>();

        for (int y = 0; y < image.Height; y++)
        {
            for (int x = 0; x < image.Width; x++)
            {
                byte alpha = image[x, y].A;
                bits.Add((alpha & 1) == 1 ? '1' : '0');

                if (CheckForTerminationSequence(bits))
                {
                    return ConvertBitsToString(bits);
                }
            }
        }

        return ConvertBitsToString(bits);
    }

    public string ExtractHiddenData(MemoryStream imageStream)
    {
        byte[] imageBytes = imageStream.ToArray();
        return ExtractHiddenData(imageBytes);
    }

    private string ConvertBitsToString(List<char> bits)
    {
        var bytes = new List<byte>();
        for (int i = 0; i < bits.Count; i += 8)
        {
            if (i + 8 > bits.Count) break;
            string byteString = new string(bits.GetRange(i, 8).ToArray());
            bytes.Add(Convert.ToByte(byteString, 2));
        }

        return Encoding.UTF8.GetString(bytes.ToArray()).Split("%%EOF%%")[0];
    }

    private bool CheckForTerminationSequence(List<char> bits)
    {
        string currentBits = new string(bits.ToArray());
        string terminationBinary = StringToBinary("%%EOF%%");
        return currentBits.Contains(terminationBinary);
    }

    private string StringToBinary(string text)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        return string.Join("", bytes.Select(b => Convert.ToString(b, 2).PadLeft(8, '0')));
    }
}

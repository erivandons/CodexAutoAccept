using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;

internal static class Program
{
    private static readonly TimeSpan Interval = TimeSpan.FromSeconds(5);
    private static readonly TimeSpan EnterCooldown = TimeSpan.FromSeconds(15);
    private static readonly TimeSpan ContinueCooldown = TimeSpan.FromSeconds(30);
    private static readonly OcrEngine? OcrEngine = Windows.Media.Ocr.OcrEngine.TryCreateFromUserProfileLanguages();

    private static DateTimeOffset _lastEnterSent = DateTimeOffset.MinValue;
    private static DateTimeOffset _lastContinueSent = DateTimeOffset.MinValue;
    private static ScreenAction _pendingAction = ScreenAction.None;
    private static int _pendingHits;

    private const int MinConfirmations = 2;

    private static readonly string[] MustContainAny =
    {
        "codex",
        "chatgpt"
    };

    private static readonly string[] ApprovalStrongWords =
    {
        "allow once", "allow always", "approve", "accept", "permission request"
    };

    private static readonly string[] ApprovalWords =
    {
        "allow", "yes", "confirm", "permission",
        "permitir", "aceitar", "sim", "confirmar", "permissao", "permissão"
    };

    private static readonly string[] NextStepWords =
    {
        "next step", "next steps", "next-step",
        "proximo passo", "próximo passo",
        "proxima etapa", "próxima etapa",
        "proximos passos", "próximos passos",
        "proximas etapas", "próximas etapas"
    };

    private static readonly string[] ContinueSuggestionWords =
    {
        "se quiser", "if you want", "i can continue", "can continue",
        "quer que eu", "posso continuar", "posso seguir", "continue daqui",
        "i can do that next", "i can proceed"
    };

    private static async Task Main()
    {
        Console.OutputEncoding = Encoding.UTF8;
        Console.WriteLine("Codex Auto Accept iniciado.");
        Console.WriteLine("Detecção reforçada: OCR em múltiplas regiões + confirmação em 2 ciclos.");
        Console.WriteLine("Pressione Ctrl+C para encerrar.\n");

        if (OcrEngine is null)
        {
            Console.WriteLine("OCR indisponível no Windows (idiomas do perfil não suportados).");
            return;
        }

        while (true)
        {
            try
            {
                IntPtr hWnd = FindVsCodeWindow();
                if (hWnd == IntPtr.Zero)
                {
                    ResetPendingAction();
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] VS Code não encontrado.");
                    await Task.Delay(Interval);
                    continue;
                }

                using Bitmap? screenshot = CaptureWindow(hWnd);
                if (screenshot is null)
                {
                    ResetPendingAction();
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Falha na captura da janela do VS Code.");
                    await Task.Delay(Interval);
                    continue;
                }

                DetectionResult detection = await AnalyzeScreenAsync(screenshot);
                if (detection.Action == ScreenAction.None)
                {
                    ResetPendingAction();
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Nenhum gatilho. Melhor score={detection.Score} ({detection.Source}).");
                    await Task.Delay(Interval);
                    continue;
                }

                bool confirmed = RegisterAndCheckConfirmation(detection.Action);
                Console.WriteLine(
                    $"[{DateTime.Now:HH:mm:ss}] Detectado {detection.Action} score={detection.Score} src={detection.Source} conf={_pendingHits}/{MinConfirmations} txt=\"{detection.Snippet}\"");

                if (!confirmed)
                {
                    await Task.Delay(Interval);
                    continue;
                }

                if (detection.Action == ScreenAction.AcceptApproval && DateTimeOffset.Now - _lastEnterSent < EnterCooldown)
                {
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Prompt detectado, aguardando cooldown.");
                    await Task.Delay(Interval);
                    continue;
                }

                if (detection.Action == ScreenAction.SendPodeSeguir && DateTimeOffset.Now - _lastContinueSent < ContinueCooldown)
                {
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Próximo passo detectado, aguardando cooldown.");
                    await Task.Delay(Interval);
                    continue;
                }

                SetForegroundWindow(hWnd);
                await Task.Delay(120);

                if (detection.Action == ScreenAction.AcceptApproval)
                {
                    SendEnter(hWnd);
                    _lastEnterSent = DateTimeOffset.Now;
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Enter enviado para VS Code.");
                }
                else
                {
                    SendTextAndEnter("pode seguir");
                    _lastContinueSent = DateTimeOffset.Now;
                    Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] 'pode seguir' enviado.");
                }

                ResetPendingAction();
            }
            catch (Exception ex)
            {
                ResetPendingAction();
                Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] Erro: {ex.Message}");
            }

            await Task.Delay(Interval);
        }
    }

    private static async Task<DetectionResult> AnalyzeScreenAsync(Bitmap screenshot)
    {
        DetectionResult best = DetectionResult.None;

        foreach ((string name, Rectangle region, bool enhanced) in EnumerateOcrTargets(screenshot.Size))
        {
            using Bitmap cropped = CropBitmap(screenshot, region);
            using Bitmap prepared = enhanced ? EnhanceForOcr(cropped) : new Bitmap(cropped);

            string rawText = await RunOcrAsync(prepared);
            if (string.IsNullOrWhiteSpace(rawText))
            {
                continue;
            }

            DetectionResult candidate = ScoreDetection(rawText, name, enhanced);
            if (candidate.Score > best.Score)
            {
                best = candidate;
            }
        }

        return best;
    }

    private static IEnumerable<(string name, Rectangle region, bool enhanced)> EnumerateOcrTargets(Size size)
    {
        Rectangle full = new(0, 0, size.Width, size.Height);
        Rectangle center = PercentRect(size, 0.18, 0.12, 0.64, 0.70);
        Rectangle rightPanel = PercentRect(size, 0.52, 0.00, 0.48, 1.00);
        Rectangle lowerHalf = PercentRect(size, 0.00, 0.45, 1.00, 0.55);
        Rectangle lowerRight = PercentRect(size, 0.45, 0.40, 0.55, 0.60);
        Rectangle bottomInput = PercentRect(size, 0.35, 0.72, 0.65, 0.26);

        yield return ("full_raw", full, false);
        yield return ("full_enh", full, true);
        yield return ("center_enh", center, true);
        yield return ("right_enh", rightPanel, true);
        yield return ("lower_enh", lowerHalf, true);
        yield return ("lower_right_enh", lowerRight, true);
        yield return ("bottom_input_enh", bottomInput, true);
    }

    private static DetectionResult ScoreDetection(string rawText, string source, bool enhanced)
    {
        string normalized = NormalizeForMatch(rawText);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return DetectionResult.None;
        }

        int codexHits = CountContains(normalized, MustContainAny);
        int codexScore = codexHits > 0 ? 4 + codexHits : 0;

        int approvalScore = (3 * CountContains(normalized, ApprovalStrongWords)) + CountContains(normalized, ApprovalWords);
        int nextScore = (2 * CountContains(normalized, NextStepWords)) + CountContains(normalized, ContinueSuggestionWords);

        bool looksLikePrompt = normalized.Contains("allow") || normalized.Contains("permitir") || normalized.Contains("accept") || normalized.Contains("aceitar");
        bool looksLikeSuggestion = normalized.Contains("next") || normalized.Contains("proxim") || normalized.Contains("seguir") || normalized.Contains("continu");

        if (approvalScore > 0)
        {
            approvalScore += looksLikePrompt ? 1 : 0;
            approvalScore += enhanced ? 1 : 0;
        }

        if (nextScore > 0)
        {
            nextScore += looksLikeSuggestion ? 1 : 0;
            nextScore += enhanced ? 1 : 0;
        }

        if (codexScore == 0)
        {
            if (approvalScore >= 6)
            {
                codexScore = 1;
            }
            else if (nextScore >= 6)
            {
                codexScore = 1;
            }
        }

        int totalApproval = codexScore + approvalScore;
        int totalNext = codexScore + nextScore;

        if (totalApproval >= 7 && totalApproval >= totalNext)
        {
            return new DetectionResult(
                ScreenAction.AcceptApproval,
                totalApproval,
                source,
                BuildSnippet(normalized));
        }

        if (totalNext >= 7)
        {
            return new DetectionResult(
                ScreenAction.SendPodeSeguir,
                totalNext,
                source,
                BuildSnippet(normalized));
        }

        return new DetectionResult(ScreenAction.None, Math.Max(totalApproval, totalNext), source, BuildSnippet(normalized));
    }

    private static bool RegisterAndCheckConfirmation(ScreenAction action)
    {
        if (_pendingAction != action)
        {
            _pendingAction = action;
            _pendingHits = 1;
            return false;
        }

        _pendingHits++;
        return _pendingHits >= MinConfirmations;
    }

    private static void ResetPendingAction()
    {
        _pendingAction = ScreenAction.None;
        _pendingHits = 0;
    }

    private static string NormalizeForMatch(string text)
    {
        string lower = text.ToLowerInvariant()
            .Replace('\r', ' ')
            .Replace('\n', ' ')
            .Replace('\t', ' ');

        lower = lower
            .Replace('0', 'o')
            .Replace('1', 'l')
            .Replace('5', 's')
            .Replace('|', 'l');

        string decomposed = lower.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder(decomposed.Length);

        foreach (char c in decomposed)
        {
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory(c);
            if (category == UnicodeCategory.NonSpacingMark)
            {
                continue;
            }

            if (char.IsLetterOrDigit(c) || c == ' ')
            {
                sb.Append(c);
            }
            else
            {
                sb.Append(' ');
            }
        }

        return CollapseSpaces(sb.ToString().Normalize(NormalizationForm.FormC));
    }

    private static string CollapseSpaces(string text)
    {
        var sb = new StringBuilder(text.Length);
        bool prevSpace = true;

        foreach (char c in text)
        {
            bool isSpace = c == ' ';
            if (isSpace)
            {
                if (!prevSpace)
                {
                    sb.Append(' ');
                }
            }
            else
            {
                sb.Append(c);
            }

            prevSpace = isSpace;
        }

        return sb.ToString().Trim();
    }

    private static int CountContains(string text, IEnumerable<string> patterns)
    {
        int count = 0;
        foreach (string pattern in patterns)
        {
            string p = NormalizeForMatch(pattern);
            if (p.Length > 0 && text.Contains(p, StringComparison.Ordinal))
            {
                count++;
            }
        }

        return count;
    }

    private static string BuildSnippet(string text)
    {
        if (text.Length <= 120)
        {
            return text;
        }

        return text[..120] + "...";
    }

    private static Rectangle PercentRect(Size size, double x, double y, double w, double h)
    {
        int left = Math.Max(0, (int)(size.Width * x));
        int top = Math.Max(0, (int)(size.Height * y));
        int width = Math.Max(1, Math.Min(size.Width - left, (int)(size.Width * w)));
        int height = Math.Max(1, Math.Min(size.Height - top, (int)(size.Height * h)));
        return new Rectangle(left, top, width, height);
    }

    private static Bitmap CropBitmap(Bitmap source, Rectangle region)
    {
        Rectangle safe = Rectangle.Intersect(new Rectangle(0, 0, source.Width, source.Height), region);
        return source.Clone(safe, PixelFormat.Format32bppArgb);
    }

    private static Bitmap EnhanceForOcr(Bitmap source)
    {
        int width = Math.Max(1, source.Width * 2);
        int height = Math.Max(1, source.Height * 2);

        var enlarged = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using (Graphics g = Graphics.FromImage(enlarged))
        {
            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
            g.DrawImage(source, new Rectangle(0, 0, width, height));
        }

        var output = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using (Graphics g = Graphics.FromImage(output))
        using (var attributes = new ImageAttributes())
        {
            const float contrast = 1.35f;
            float t = 0.5f * (1f - contrast);

            var matrix = new ColorMatrix(new[]
            {
                new[] { 0.2126f * contrast, 0.2126f * contrast, 0.2126f * contrast, 0f, 0f },
                new[] { 0.7152f * contrast, 0.7152f * contrast, 0.7152f * contrast, 0f, 0f },
                new[] { 0.0722f * contrast, 0.0722f * contrast, 0.0722f * contrast, 0f, 0f },
                new[] { 0f, 0f, 0f, 1f, 0f },
                new[] { t, t, t, 0f, 1f }
            });

            attributes.SetColorMatrix(matrix);
            g.DrawImage(
                enlarged,
                new Rectangle(0, 0, width, height),
                0,
                0,
                width,
                height,
                GraphicsUnit.Pixel,
                attributes);
        }

        enlarged.Dispose();
        return output;
    }

    private static async Task<string> RunOcrAsync(Bitmap bitmap)
    {
        if (OcrEngine is null)
        {
            return string.Empty;
        }

        using var ms = new MemoryStream();
        bitmap.Save(ms, ImageFormat.Png);
        ms.Position = 0;

        using var ras = new InMemoryRandomAccessStream();
        byte[] bytes = ms.ToArray();

        using (var writer = new DataWriter(ras.GetOutputStreamAt(0)))
        {
            writer.WriteBytes(bytes);
            await writer.StoreAsync();
            await writer.FlushAsync();
        }

        ras.Seek(0);
        BitmapDecoder decoder = await BitmapDecoder.CreateAsync(ras);
        SoftwareBitmap softwareBitmap = await decoder.GetSoftwareBitmapAsync();

        OcrResult result = await OcrEngine.RecognizeAsync(softwareBitmap);
        return result.Text ?? string.Empty;
    }

    private static IntPtr FindVsCodeWindow()
    {
        IntPtr foreground = GetForegroundWindow();
        if (foreground != IntPtr.Zero && IsVsCodeWindow(foreground))
        {
            return foreground;
        }

        foreach (Process process in Process.GetProcessesByName("Code"))
        {
            if (process.MainWindowHandle != IntPtr.Zero)
            {
                return process.MainWindowHandle;
            }
        }

        return IntPtr.Zero;
    }

    private static bool IsVsCodeWindow(IntPtr hWnd)
    {
        GetWindowThreadProcessId(hWnd, out uint pid);

        try
        {
            Process process = Process.GetProcessById((int)pid);
            return process.ProcessName.Equals("Code", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static Bitmap? CaptureWindow(IntPtr hWnd)
    {
        if (!GetWindowRect(hWnd, out RECT rect))
        {
            return null;
        }

        int width = rect.Right - rect.Left;
        int height = rect.Bottom - rect.Top;
        if (width <= 0 || height <= 0)
        {
            return null;
        }

        var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using Graphics g = Graphics.FromImage(bmp);
        g.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
        return bmp;
    }

    private static void SendEnter(IntPtr hWnd)
    {
        PostMessage(hWnd, WM_KEYDOWN, (IntPtr)VK_RETURN, IntPtr.Zero);
        PostMessage(hWnd, WM_KEYUP, (IntPtr)VK_RETURN, IntPtr.Zero);
    }

    private static void SendTextAndEnter(string text)
    {
        foreach (char ch in text)
        {
            SendUnicodeChar(ch);
        }

        SendVirtualKey(VK_RETURN);
    }

    private static void SendUnicodeChar(char ch)
    {
        INPUT[] inputs =
        {
            new()
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wScan = ch,
                        dwFlags = KEYEVENTF_UNICODE
                    }
                }
            },
            new()
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wScan = ch,
                        dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP
                    }
                }
            }
        };

        SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
    }

    private static void SendVirtualKey(ushort virtualKey)
    {
        INPUT[] inputs =
        {
            new()
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = virtualKey
                    }
                }
            },
            new()
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = virtualKey,
                        dwFlags = KEYEVENTF_KEYUP
                    }
                }
            }
        };

        SendInput((uint)inputs.Length, inputs, Marshal.SizeOf<INPUT>());
    }

    private readonly record struct DetectionResult(ScreenAction Action, int Score, string Source, string Snippet)
    {
        public static DetectionResult None => new(ScreenAction.None, 0, "-", "");
    }

    private enum ScreenAction
    {
        None,
        AcceptApproval,
        SendPodeSeguir
    }

    private const uint WM_KEYDOWN = 0x0100;
    private const uint WM_KEYUP = 0x0101;
    private const ushort VK_RETURN = 0x0D;
    private const uint INPUT_KEYBOARD = 1;
    private const uint KEYEVENTF_KEYUP = 0x0002;
    private const uint KEYEVENTF_UNICODE = 0x0004;

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint type;
        public InputUnion U;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)]
        public KEYBDINPUT ki;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KEYBDINPUT
    {
        public ushort wVk;
        public ushort wScan;
        public uint dwFlags;
        public uint time;
        public IntPtr dwExtraInfo;
    }
}

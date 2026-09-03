using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace WhackAMole;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new GameForm());
    }
}

public sealed class GameForm : Form
{
    private readonly Button[] holes = new Button[9];
    private readonly Label scoreLabel = new();
    private readonly Label timeLabel = new();
    private readonly System.Windows.Forms.Timer moleTimer = new() { Interval = 650 };
    private readonly System.Windows.Forms.Timer gameTimer = new() { Interval = 1000 };
    private readonly Random random = new();
    private int activeHole = -1;
    private int score;
    private int secondsLeft = 30;

    public GameForm()
    {
        Text = "Whack-a-Mole - C# WinForms";
        ClientSize = new Size(520, 600);
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;

        scoreLabel.Text = "Score: 0";
        scoreLabel.Font = new Font("Segoe UI", 16, FontStyle.Bold);
        scoreLabel.Location = new Point(25, 20);
        scoreLabel.AutoSize = true;
        Controls.Add(scoreLabel);

        timeLabel.Text = "Time: 30";
        timeLabel.Font = new Font("Segoe UI", 16, FontStyle.Bold);
        timeLabel.Location = new Point(370, 20);
        timeLabel.AutoSize = true;
        Controls.Add(timeLabel);

        var grid = new TableLayoutPanel { Location = new Point(35, 80), Size = new Size(450, 450), RowCount = 3, ColumnCount = 3 };
        for (int i = 0; i < 3; i++)
        {
            grid.RowStyles.Add(new RowStyle(SizeType.Percent, 33.333f));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.333f));
        }

        for (int i = 0; i < holes.Length; i++)
        {
            var button = new Button { Dock = DockStyle.Fill, Margin = new Padding(7), Font = new Font("Segoe UI Emoji", 32), Text = "○", Tag = i };
            button.Click += HoleClicked;
            holes[i] = button;
            grid.Controls.Add(button, i % 3, i / 3);
        }
        Controls.Add(grid);

        var hint = new Label { Text = "30 秒内点击出现的 🐹", Font = new Font("Segoe UI", 11), AutoSize = true, Location = new Point(175, 550) };
        Controls.Add(hint);

        moleTimer.Tick += (_, _) => ShowMole();
        gameTimer.Tick += (_, _) => TickGame();
        FormClosing += (_, _) => SaveDiagnostics();
        moleTimer.Start();
        gameTimer.Start();
        ShowMole();
    }

    private void ShowMole()
    {
        if (activeHole >= 0) holes[activeHole].Text = "○";
        activeHole = random.Next(holes.Length);
        holes[activeHole].Text = "🐹";
    }

    private void HoleClicked(object? sender, EventArgs e)
    {
        if (sender is Button button && (int)button.Tag! == activeHole && secondsLeft > 0)
        {
            score++;
            scoreLabel.Text = $"Score: {score}";
            holes[activeHole].Text = "✓";
            activeHole = -1;
        }
    }

    private void TickGame()
    {
        secondsLeft--;
        timeLabel.Text = $"Time: {secondsLeft}";
        if (secondsLeft <= 0)
        {
            moleTimer.Stop();
            gameTimer.Stop();
            if (activeHole >= 0) holes[activeHole].Text = "○";
            activeHole = -1;
            MessageBox.Show($"时间到！你的分数：{score}", "Game Over");
        }
    }

    private void SaveDiagnostics()
    {
        try
        {
            Directory.CreateDirectory("output");
            var text = new StringBuilder()
                .AppendLine("game=C# Whack-a-Mole")
                .AppendLine($"time={DateTime.Now:O}")
                .AppendLine($"score={score}")
                .AppendLine($"os={RuntimeInformation.OSDescription}")
                .AppendLine($"framework={RuntimeInformation.FrameworkDescription}")
                .AppendLine($"arch={RuntimeInformation.ProcessArchitecture}")
                .ToString();
            File.WriteAllText(Path.Combine("output", "csharp_whackamole_last_run.txt"), text, Encoding.UTF8);
        }
        catch { }
    }
}

using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace AlerteSalaries
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            bool configureMode = false;
            foreach (string arg in args)
            {
                if (string.Equals(arg, "/configure", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(arg, "-configure", StringComparison.OrdinalIgnoreCase))
                {
                    configureMode = true;
                }
            }

            AppRuntime runtime = new AppRuntime();
            runtime.Run(configureMode);
        }
    }

    internal sealed class AppRuntime
    {
        private const int RecordingDurationSeconds = 900;
        private const string AppName = "AlerteSalaries";
        private const string ActiveAlertMutexName = @"Local\AlerteSalaries.ActiveAlert";
        private const string LegacyConfigFileName = "config.json";
        private const string UserConfigFileName = "user.config.json";
        private const string AdminConfigFileName = "AlerteSalaries.Admin.ini";

        public string AppDataDirectory { get; private set; }
        public string UserConfigPath { get; private set; }
        public string LegacyConfigPath { get; private set; }
        public string CommonDataDirectory { get; private set; }
        public string AdminConfigPath { get; private set; }
        public string LogPath { get; private set; }

        public AppRuntime()
        {
            AppDataDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                AppName);
            CommonDataDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                AppName);
            UserConfigPath = Path.Combine(AppDataDirectory, UserConfigFileName);
            LegacyConfigPath = Path.Combine(AppDataDirectory, LegacyConfigFileName);
            AdminConfigPath = Path.Combine(CommonDataDirectory, AdminConfigFileName);
            LogPath = Path.Combine(AppDataDirectory, "app.log");
            Directory.CreateDirectory(AppDataDirectory);
            Directory.CreateDirectory(CommonDataDirectory);
        }

        public void Run(bool configureMode)
        {
            UserConfig userConfig = LoadUserConfig();
            bool initialSetup = userConfig == null;

            if (configureMode || initialSetup)
            {
                using (ConfigurationForm form = new ConfigurationForm(userConfig))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        SaveUserConfig(form.Result);
                        WriteLog("Configuration enregistree.");
                        userConfig = form.Result;
                    }
                    else
                    {
                        WriteLog("Configuration annulee.");
                    }
                }

                if (configureMode || initialSetup)
                {
                    return;
                }
            }

            bool createdNew;
            using (Mutex mutex = new Mutex(false, ActiveAlertMutexName, out createdNew))
            {
                if (!createdNew)
                {
                    WriteLog("Tentative ignoree: une alerte est deja en cours.");
                    return;
                }

                try
                {
                    AdminConfig adminConfig = LoadAdminConfig();
                    if (adminConfig != null && !string.IsNullOrWhiteSpace(adminConfig.WebhookUrl))
                    {
                        SendTeamsAlert(userConfig, adminConfig);
                    }
                    else
                    {
                        WriteLog("Alerte Teams non envoyee: configuration admin absente ou incomplete.");
                    }
                }
                catch (Exception ex)
                {
                    WriteLog("Echec envoi Teams: " + ex.Message);
                }

                try
                {
                    RecordAudio(userConfig, RecordingDurationSeconds);
                }
                catch (Exception ex)
                {
                    WriteLog("Echec enregistrement audio: " + ex.Message);
                }
            }
        }

        private UserConfig LoadUserConfig()
        {
            if (File.Exists(UserConfigPath))
            {
                try
                {
                    using (FileStream stream = File.OpenRead(UserConfigPath))
                    {
                        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(UserConfig));
                        return (UserConfig)serializer.ReadObject(stream);
                    }
                }
                catch (Exception ex)
                {
                    WriteLog("Erreur de lecture de configuration utilisateur: " + ex.Message);
                    return null;
                }
            }

            if (File.Exists(LegacyConfigPath))
            {
                try
                {
                    using (FileStream stream = File.OpenRead(LegacyConfigPath))
                    {
                        DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(LegacyUserConfig));
                        LegacyUserConfig legacyConfig = (LegacyUserConfig)serializer.ReadObject(stream);
                        if (legacyConfig == null)
                        {
                            return null;
                        }

                        UserConfig migratedConfig = new UserConfig
                        {
                            EmployeeName = legacyConfig.EmployeeName,
                            OfficeLocations = legacyConfig.OfficeLocations
                        };
                        SaveUserConfig(migratedConfig);
                        WriteLog("Configuration utilisateur migree depuis l'ancien format.");
                        return migratedConfig;
                    }
                }
                catch (Exception ex)
                {
                    WriteLog("Erreur de lecture de l'ancienne configuration utilisateur: " + ex.Message);
                    return null;
                }
            }

            return null;
        }

        private void SaveUserConfig(UserConfig config)
        {
            using (FileStream stream = File.Create(UserConfigPath))
            {
                DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(UserConfig));
                serializer.WriteObject(stream, config);
            }
        }

        private AdminConfig LoadAdminConfig()
        {
            string[] candidatePaths = new[]
            {
                AdminConfigPath,
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, AdminConfigFileName)
            };

            foreach (string path in candidatePaths)
            {
                if (!File.Exists(path))
                {
                    continue;
                }

                try
                {
                    return ParseAdminConfig(path);
                }
                catch (Exception ex)
                {
                    WriteLog("Erreur de lecture de configuration admin '" + path + "': " + ex.Message);
                }
            }

            return null;
        }

        private static AdminConfig ParseAdminConfig(string path)
        {
            AdminConfig config = new AdminConfig();
            string[] lines = File.ReadAllLines(path, Encoding.UTF8);

            foreach (string rawLine in lines)
            {
                string line = rawLine.Trim();
                if (line.Length == 0 || line.StartsWith(";") || line.StartsWith("#") || line.StartsWith("["))
                {
                    continue;
                }

                int separatorIndex = line.IndexOf('=');
                if (separatorIndex <= 0)
                {
                    continue;
                }

                string key = line.Substring(0, separatorIndex).Trim();
                string value = line.Substring(separatorIndex + 1).Trim();

                if (string.Equals(key, "WebhookUrl", StringComparison.OrdinalIgnoreCase))
                {
                    config.WebhookUrl = value;
                }
            }

            return config;
        }

        private void SendTeamsAlert(UserConfig userConfig, AdminConfig adminConfig)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            string machineName = Environment.MachineName;
            string currentUser = WindowsIdentity.GetCurrent().Name;
            string timestamp = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            string payload = BuildTeamsAdaptiveCardPayload(
                "ALERTE SECURITE",
                userConfig.EmployeeName,
                userConfig.OfficeLocations,
                machineName,
                currentUser,
                timestamp);
            byte[] bytes = Encoding.UTF8.GetBytes(payload);

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(adminConfig.WebhookUrl);
            request.Method = "POST";
            request.ContentType = "application/json";
            request.ContentLength = bytes.Length;

            using (Stream requestStream = request.GetRequestStream())
            {
                requestStream.Write(bytes, 0, bytes.Length);
            }

            using ((HttpWebResponse)request.GetResponse())
            {
            }

            WriteLog("Alerte Teams envoyee pour " + userConfig.EmployeeName + ".");
        }

        private static string BuildTeamsAdaptiveCardPayload(
            string title,
            string employeeName,
            string officeLocations,
            string machineName,
            string currentUser,
            string timestamp)
        {
            return "{"
                + "\"type\":\"message\","
                + "\"attachments\":[{"
                + "\"contentType\":\"application/vnd.microsoft.card.adaptive\","
                + "\"contentUrl\":null,"
                + "\"content\":{"
                + "\"$schema\":\"http://adaptivecards.io/schemas/adaptive-card.json\","
                + "\"type\":\"AdaptiveCard\","
                + "\"version\":\"1.4\","
                + "\"msteams\":{\"width\":\"Full\"},"
                + "\"body\":["
                + "{"
                + "\"type\":\"TextBlock\","
                + "\"text\":\"" + JsonEscape(title) + "\","
                + "\"weight\":\"Bolder\","
                + "\"size\":\"Large\","
                + "\"color\":\"Attention\","
                + "\"wrap\":true"
                + "},"
                + "{"
                + "\"type\":\"TextBlock\","
                + "\"text\":\"Une assistance immediate est requise.\","
                + "\"wrap\":true,"
                + "\"spacing\":\"Small\""
                + "},"
                + "{"
                + "\"type\":\"FactSet\","
                + "\"facts\":["
                + "{\"title\":\"Collaborateur\",\"value\":\"" + JsonEscape(employeeName) + "\"},"
                + "{\"title\":\"Bureau habituel\",\"value\":\"" + JsonEscape(officeLocations) + "\"},"
                + "{\"title\":\"Poste\",\"value\":\"" + JsonEscape(machineName) + "\"},"
                + "{\"title\":\"Session\",\"value\":\"" + JsonEscape(currentUser) + "\"},"
                + "{\"title\":\"Heure\",\"value\":\"" + JsonEscape(timestamp) + "\"}"
                + "]"
                + "}"
                + "]"
                + "}"
                + "}]"
                + "}";
        }

        private void RecordAudio(UserConfig config, int durationSeconds)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string safeName = MakeSafeFileName(config.EmployeeName);
            string timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
            string baseName = "alerte-" + safeName + "-" + timestamp;
            string wavPath = Path.Combine(desktopPath, baseName + ".wav");
            string mp3Path = Path.Combine(desktopPath, baseName + ".mp3");

            WriteLog("Enregistrement audio demarre.");
            NativeMethods.ThrowIfMciError(NativeMethods.MciSendString("open new Type waveaudio Alias alertrec"));
            try
            {
                NativeMethods.ThrowIfMciError(NativeMethods.MciSendString("record alertrec"));
                Thread.Sleep(TimeSpan.FromSeconds(durationSeconds));
                NativeMethods.ThrowIfMciError(NativeMethods.MciSendString("stop alertrec"));
                NativeMethods.ThrowIfMciError(NativeMethods.MciSendString("save alertrec \"" + wavPath + "\""));
            }
            finally
            {
                try
                {
                    NativeMethods.MciSendString("close alertrec");
                }
                catch
                {
                }
            }

            string ffmpegPath = FindFfmpeg();
            if (!string.IsNullOrEmpty(ffmpegPath))
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = ffmpegPath,
                    Arguments = "-y -i \"" + wavPath + "\" -codec:a libmp3lame -q:a 4 \"" + mp3Path + "\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                using (Process process = Process.Start(startInfo))
                {
                    process.WaitForExit();
                    if (process.ExitCode == 0 && File.Exists(mp3Path))
                    {
                        File.Delete(wavPath);
                        WriteLog("Audio converti en mp3: " + mp3Path);
                        return;
                    }
                }
            }

            WriteLog("Audio sauvegarde en wav: " + wavPath);
        }

        private string FindFfmpeg()
        {
            string localFfmpeg = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ffmpeg.exe");
            if (File.Exists(localFfmpeg))
            {
                return localFfmpeg;
            }

            string path = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            string[] directories = path.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string directory in directories)
            {
                try
                {
                    string candidate = Path.Combine(directory.Trim(), "ffmpeg.exe");
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        private static string MakeSafeFileName(string value)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                value = value.Replace(c, '_');
            }

            return value;
        }

        private static string JsonEscape(string value)
        {
            return value
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r", "\\r")
                .Replace("\n", "\\n");
        }

        private void WriteLog(string message)
        {
            string line = "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] " + message + Environment.NewLine;
            File.AppendAllText(LogPath, line, Encoding.UTF8);
        }
    }

    [DataContract]
    internal sealed class UserConfig
    {
        [DataMember(Name = "employeeName")]
        public string EmployeeName { get; set; }

        [DataMember(Name = "officeLocations")]
        public string OfficeLocations { get; set; }
    }

    [DataContract]
    internal sealed class LegacyUserConfig
    {
        [DataMember(Name = "employeeName")]
        public string EmployeeName { get; set; }

        [DataMember(Name = "officeLocations")]
        public string OfficeLocations { get; set; }
    }

    internal sealed class AdminConfig
    {
        public string WebhookUrl { get; set; }
    }

    internal sealed class ConfigurationForm : Form
    {
        private readonly TextBox employeeNameTextBox;
        private readonly TextBox officeLocationsTextBox;
        private readonly Font titleFont;
        private readonly Font sectionFont;

        public UserConfig Result { get; private set; }

        public ConfigurationForm(UserConfig config)
        {
            titleFont = new Font("Segoe UI", 16F, FontStyle.Bold);
            sectionFont = new Font("Segoe UI", 9.5F, FontStyle.Bold);

            Text = "Alerte Salaries - Configuration initiale";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            TopMost = true;
            BackColor = Color.White;
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScroll = true;
            ClientSize = new Size(820, 760);

            try
            {
                Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
            }
            catch
            {
            }

            Panel headerPanel = new Panel
            {
                Left = 0,
                Top = 0,
                Width = 820,
                Height = 92,
                BackColor = Color.FromArgb(255, 247, 195, 37)
            };
            Controls.Add(headerPanel);

            Label iconLabel = new Label
            {
                Left = 18,
                Top = 18,
                Width = 70,
                Height = 52,
                Font = new Font("Segoe UI Symbol", 28F, FontStyle.Bold),
                Text = "⚠",
                ForeColor = Color.FromArgb(255, 36, 36, 36)
            };
            headerPanel.Controls.Add(iconLabel);

            headerPanel.Controls.Add(new Label
            {
                Left = 86,
                Top = 18,
                Width = 640,
                Height = 28,
                Font = titleFont,
                Text = "Alerte Salaries"
            });

            headerPanel.Controls.Add(new Label
            {
                Left = 88,
                Top = 48,
                Width = 640,
                Height = 26,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular),
                Text = "Bouton d'alerte discret pour demander de l'aide et enregistrer le micro pendant 15 minutes."
            });

            int contentLeft = 22;
            int contentWidth = 760;
            int inputLeft = 250;
            int inputWidth = 520;
            int y = 112;

            Label section1 = CreateSectionLabel("A quoi sert cette application ?", contentLeft, y, contentWidth);
            Controls.Add(section1);
            y = section1.Bottom + 8;

            Label body1 = CreateWrappingLabel(
                "Quand l'utilisateur clique sur l'icone Alerte Salaries, aucun ecran normal ne s'ouvre. "
                + "L'application envoie une alerte au canal Teams configure, puis enregistre le micro du poste pendant 15 minutes. "
                + "Le fichier audio est depose sur le Bureau du poste concerne.",
                contentLeft, y, contentWidth, Color.Black);
            Controls.Add(body1);
            y = body1.Bottom + 22;

            Label section2 = CreateSectionLabel("Que faire sur cet ecran ?", contentLeft, y, contentWidth);
            Controls.Add(section2);
            y = section2.Bottom + 8;

            Label body2 = CreateWrappingLabel(
                "Remplissez les champs ci-dessous une seule fois. Au prochain clic sur l'icone, l'alerte sera envoyee automatiquement. "
                + "Si vous devez modifier la configuration plus tard, utilisez le raccourci 'Configurer Alerte Salaries'.",
                contentLeft, y, contentWidth, Color.Black);
            Controls.Add(body2);
            y = body2.Bottom + 22;

            Label section3 = CreateSectionLabel("Information geree par votre administrateur", contentLeft, y, contentWidth);
            Controls.Add(section3);
            y = section3.Bottom + 8;

            Label body3 = CreateWrappingLabel(
                "Le lien Teams de reception d'alerte est configure automatiquement par l'administration lors du deploiement. "
                + "Vous n'avez rien a saisir pour cette partie.",
                contentLeft, y, contentWidth, Color.Black);
            Controls.Add(body3);
            y = body3.Bottom + 22;

            Label section4 = CreateSectionLabel("Informations a renseigner", contentLeft, y, contentWidth);
            Controls.Add(section4);
            y = section4.Bottom + 16;

            Label employeeLabel = CreateLabel("Nom du collaborateur", contentLeft, y + 4);
            Controls.Add(employeeLabel);
            employeeNameTextBox = CreateTextBox(config != null ? config.EmployeeName : Environment.UserName, inputLeft, y, inputWidth);
            Controls.Add(employeeNameTextBox);
            y = employeeNameTextBox.Bottom + 6;

            Label employeeHint = CreateWrappingLabel("Nom qui apparaitra dans le message d'alerte Teams.", inputLeft, y, inputWidth, Color.FromArgb(255, 90, 90, 90));
            Controls.Add(employeeHint);
            y = employeeHint.Bottom + 18;

            Label officeLabel = CreateLabel("Bureau habituel", contentLeft, y + 4);
            Controls.Add(officeLabel);
            officeLocationsTextBox = CreateTextBox(config != null ? config.OfficeLocations : string.Empty, inputLeft, y, inputWidth);
            Controls.Add(officeLocationsTextBox);
            y = officeLocationsTextBox.Bottom + 6;

            Label officeHint = CreateWrappingLabel(
                "Exemple: Accueil Paris, Bureau 2e etage, Site Lyon. Vous pouvez mettre plusieurs lieux separes par des virgules.",
                inputLeft,
                y,
                inputWidth,
                Color.FromArgb(255, 90, 90, 90));
            Controls.Add(officeHint);
            y = officeHint.Bottom + 18;

            Panel helpPanel = new Panel
            {
                Left = contentLeft,
                Top = y,
                Width = contentWidth,
                Height = 60,
                BackColor = Color.FromArgb(255, 247, 248, 250)
            };
            Controls.Add(helpPanel);

            Label helpLabel = CreateWrappingLabel(
                "En cas de doute sur le bureau habituel a indiquer, mettez le lieu ou vous travaillez le plus souvent ou le site qui aidera vos collegues a vous localiser rapidement.",
                12,
                8,
                730,
                Color.Black);
            helpPanel.Controls.Add(helpLabel);
            helpPanel.Height = helpLabel.Bottom + 10;
            y = helpPanel.Bottom + 26;

            Button saveButton = new Button
            {
                Text = "Enregistrer",
                Left = 612,
                Top = y,
                Width = 90,
                Height = 30
            };
            saveButton.Click += SaveButtonOnClick;
            Controls.Add(saveButton);

            Button cancelButton = new Button
            {
                Text = "Annuler",
                Left = 712,
                Top = y,
                Width = 90,
                Height = 30
            };
            cancelButton.Click += delegate
            {
                DialogResult = DialogResult.Cancel;
                Close();
            };
            Controls.Add(cancelButton);

            AcceptButton = saveButton;
            CancelButton = cancelButton;
            AutoScrollMinSize = new Size(0, cancelButton.Bottom + 24);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (titleFont != null)
                {
                    titleFont.Dispose();
                }

                if (sectionFont != null)
                {
                    sectionFont.Dispose();
                }
            }

            base.Dispose(disposing);
        }

        private void SaveButtonOnClick(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(employeeNameTextBox.Text) ||
                string.IsNullOrWhiteSpace(officeLocationsTextBox.Text))
            {
                MessageBox.Show(
                    "Tous les champs sont obligatoires.",
                    "Configuration incomplete",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            Result = new UserConfig
            {
                EmployeeName = employeeNameTextBox.Text.Trim(),
                OfficeLocations = officeLocationsTextBox.Text.Trim()
            };

            DialogResult = DialogResult.OK;
            Close();
        }

        private static Label CreateLabel(string text, int left, int top)
        {
            return new Label
            {
                Text = text,
                Left = left,
                Top = top,
                Width = 210
            };
        }

        private static TextBox CreateTextBox(string text, int left, int top, int width)
        {
            return new TextBox
            {
                Text = text,
                Left = left,
                Top = top,
                Width = width
            };
        }

        private Label CreateSectionLabel(string text, int left, int top, int width)
        {
            return new Label
            {
                Text = text,
                Left = left,
                Top = top,
                Width = width,
                AutoSize = true,
                Font = sectionFont
            };
        }

        private static Label CreateWrappingLabel(string text, int left, int top, int width, Color color)
        {
            return new Label
            {
                Text = text,
                Left = left,
                Top = top,
                MaximumSize = new Size(width, 0),
                AutoSize = true,
                ForeColor = color
            };
        }
    }

    internal static class NativeMethods
    {
        [DllImport("winmm.dll", CharSet = CharSet.Auto)]
        private static extern int mciSendString(string command, StringBuilder returnValue, int returnLength, IntPtr winHandle);

        public static int MciSendString(string command)
        {
            StringBuilder builder = new StringBuilder(255);
            return mciSendString(command, builder, builder.Capacity, IntPtr.Zero);
        }

        public static void ThrowIfMciError(int code)
        {
            if (code != 0)
            {
                throw new InvalidOperationException("Commande MCI en echec. Code: " + code);
            }
        }
    }
}

using System;
using System.Data.SqlClient;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;



namespace ProyectoFinal1
{
    public partial class Form1 : Form
    {
        private readonly HttpClient client = new HttpClient();
        private const string connectionString = "Server=localhost; Database=ProyectoFinal1;Trusted_Connection=True";
       
        private const string openAIApiUrl = "https://api.openai.com/v1/chat/completions";
        private const string openAIApiKey = "Api_key_real"; // ¡REEMPLAZA ESTO CON TU CLAVE REAL!

        private string researchResult = "";

        public Form1()
        {
            InitializeComponent();
            BtnResearch.Click += BtnResearch_Click;
            BtnApprove.Click += BtnApprove_Click;
        }

        private bool TestDatabaseConnection()
        {
            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    return true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error de conexión a la base de datos: {ex.Message}", "Error de Conexión", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private async void BtnResearch_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(TxtPrompt.Text))
            {
                MessageBox.Show("Por favor, introduce un tema de investigación.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!TestDatabaseConnection())
            {
                LblStatus.Text = "No se pudo conectar a la base de datos.";
                return;
            }

            LblStatus.Text = "Investigando...";
            TxtResult.Text = "Cargando...";

            try
            {
                if (!string.IsNullOrEmpty(TxtPrompt.Text))
                {
                    researchResult = await CallOpenAIApi(TxtPrompt.Text); // Llama a la API de OpenAI
                    TxtResult.Text = researchResult;
                    SaveToDatabase(TxtPrompt.Text, researchResult);
                    LblStatus.Text = "Investigación completada y guardada en la base de datos.";
                }
                else
                {
                    LblStatus.Text = "El prompt no puede estar vacío.";
                    TxtResult.Text = "";
                }
            }
            catch (HttpRequestException ex) // Excepción común para errores HTTP
            {
                MessageBox.Show($"Error de API de OpenAI: {ex.Message}\nVerifica tu clave API y la URL del endpoint.", "Error de API", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LblStatus.Text = "Ocurrió un error con la API de OpenAI.";
                TxtResult.Text = $"Error de API: {ex.Message}";
            }
            catch (JsonException ex) // Si la respuesta de la API no es JSON válido
            {
                MessageBox.Show($"Error al procesar la respuesta de la API: {ex.Message}", "Error de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LblStatus.Text = "Error al procesar la respuesta de la API.";
                TxtResult.Text = $"Error de JSON: {ex.Message}";
            }
            catch (SqlException ex)
            {
                MessageBox.Show($"Error de base de datos: {ex.Message}\nVerifica tu cadena de conexión y la tabla 'ResearchResults'.", "Error de Base de Datos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LblStatus.Text = "Ocurrió un error de base de datos.";
                TxtResult.Text = $"Error de DB: {ex.Message}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inesperado: {ex.Message}", "Error General", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LblStatus.Text = "Ocurrió un error inesperado.";
                TxtResult.Text = $"Error: {ex.Message}";
            }
        }

        // Método para llamar a la API de OpenAI (GPT-3.5-Turbo)
        private async Task<string> CallOpenAIApi(string prompt)
        {
            client.DefaultRequestHeaders.Clear();
            client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", openAIApiKey);

            var requestBody = new
            {
                model = "gpt-3.5-turbo", // Modelo de OpenAI para chat completions
                messages = new[] 
                {
                    new { role = "user", content = prompt }
                },
                max_tokens = 500, // Límite de tokens en la respuesta
                temperature = 0.7 // Controla la creatividad, más alto es más creativo)
            };

            var content = new StringContent(JsonSerializer.Serialize(requestBody), Encoding.UTF8, "application/json");

            var response = await client.PostAsync(openAIApiUrl, content);
            response.EnsureSuccessStatusCode(); // Lanza una excepción si la respuesta no es 2xx
            var responseBody = await response.Content.ReadAsStringAsync();
            var jsonDoc = JsonDocument.Parse(responseBody);

    
            if (jsonDoc.RootElement.TryGetProperty("choices", out var choices) && choices.GetArrayLength() > 0)
            {
                if (choices[0].TryGetProperty("message", out var message) && message.TryGetProperty("content", out var contentProperty))
                {
                    return contentProperty.GetString();
                }
            }
            return "No se pudo generar contenido con OpenAI.";
        }

        private void SaveToDatabase(string prompt, string result)
        {
            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
                var query = "INSERT INTO ResearchResults (Prompt, Result, CreatedAt) VALUES (@Prompt, @Result, @CreatedAt)";
                using (var cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@Prompt", prompt);
                    cmd.Parameters.AddWithValue("@Result", result);
                    cmd.Parameters.AddWithValue("@CreatedAt", DateTime.Now);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void BtnApprove_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(researchResult))
            {
                MessageBox.Show("No hay resultados de investigación para generar documentos.", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                string outputDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ResearchOutput");
                Directory.CreateDirectory(outputDir);
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string wordFile = Path.Combine(outputDir, $"ReporteDeInvestigacion_{timestamp}.docx");
                string pptFile = Path.Combine(outputDir, $"PresentacionDeInvestigacion_{timestamp}.pptx");

                GenerateWordDocument(wordFile, TxtPrompt.Text, researchResult);
                GeneratePowerPoint(pptFile, TxtPrompt.Text, researchResult);

                LblStatus.Text = $"Documentos generados y guardados en {outputDir}";
                MessageBox.Show($"Documentos guardados en:\n{outputDir}", "Documentos Generados", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al generar documentos: {ex.Message}\nAsegúrate de tener Microsoft Office instalado y las referencias Interop correctas (versión 16.0 para Office 2019).", "Error de Documentos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LblStatus.Text = "Ocurrió un error al generar documentos.";
            }
        }

        private void GenerateWordDocument(string filePath, string prompt, string result)
        {
            Microsoft.Office.Interop.Word.Application wordApp = null;
            Document doc = null;
            try
            {
                wordApp = new Microsoft.Office.Interop.Word.Application();
                wordApp.Visible = false;
                doc = wordApp.Documents.Add();

                var title = doc.Paragraphs.Add();
                title.Range.Text = "Informe de Investigación\n";
                title.Range.Font.Size = 16;
                title.Range.Font.Bold = 1;
                title.Range.ParagraphFormat.SpaceAfter = 12;
                title.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                var promptPara = doc.Paragraphs.Add();
                promptPara.Range.Text = $"Pregunta de Investigación: {prompt}\n\n";
                promptPara.Range.Font.Size = 12;
                promptPara.Range.Font.Bold = 1;

                var resultPara = doc.Paragraphs.Add();
                resultPara.Range.Text = $"Resultados:\n{result}\n";
                resultPara.Range.Font.Size = 12;
                resultPara.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;

                doc.SaveAs2(filePath);
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                    Marshal.ReleaseComObject(doc);
                }
                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }
            }
        }

        private void GeneratePowerPoint(string filePath, string prompt, string result)
        {
            Microsoft.Office.Interop.PowerPoint.Application pptApp = null;
            Presentation pres = null;
            try
            {
                pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
                pptApp.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                pres = pptApp.Presentations.Add();
                var slideLayout = PpSlideLayout.ppLayoutText;

                // Diapositiva 1: Título
                var slide1 = pres.Slides.Add(1, PpSlideLayout.ppLayoutTitleOnly);
                slide1.Shapes.Title.TextFrame.TextRange.Text = "Presentación de Investigación";
                slide1.Shapes.Title.TextFrame.TextRange.Font.Size = 44;
                slide1.Shapes.Title.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;

                // Diapositiva 2: Pregunta
                var slide2 = pres.Slides.Add(2, slideLayout);
                slide2.Shapes.Title.TextFrame.TextRange.Text = "Pregunta de Investigación";
                slide2.Shapes[2].TextFrame.TextRange.Text = prompt;
                slide2.Shapes[2].TextFrame.TextRange.Font.Size = 28;

                // Diapositiva 3: Resultados
                var slide3 = pres.Slides.Add(3, slideLayout);
                slide3.Shapes.Title.TextFrame.TextRange.Text = "Resultados";
                slide3.Shapes[2].TextFrame.TextRange.Text = result;
                slide3.Shapes[2].TextFrame.TextRange.Font.Size = 20;

                pres.SaveAs(filePath, PpSaveAsFileType.ppSaveAsDefault);
            }
            finally
            {
                if (pres != null)
                {
                    pres.Close();
                    Marshal.ReleaseComObject(pres);
                }
                if (pptApp != null)
                {
                    pptApp.Quit();
                    Marshal.ReleaseComObject(pptApp);
                }
            }
        }

       // private void TxtPrompt_TextChanged(object sender, EventArgs e) { }
        //private void TxtResult_TextChanged(object sender, EventArgs e) { }
    }
}
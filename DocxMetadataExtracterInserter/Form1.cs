using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using MongoDB.Bson;
using MongoDB.Driver;

namespace DocxMetadataExtracterInserter
{
    public partial class Form1 : Form
    {
        static IMongoClient _client;
        static IMongoDatabase _database;
        public Form1()
        {
            InitializeComponent();
            _client = new MongoClient();
            _database = _client.GetDatabase("CBIRtest");
        }

        private void chooseFolder_Click(object sender, EventArgs e)
        {
            DialogResult res = folderBrowserDialog1.ShowDialog();
            if (res == DialogResult.OK)
            {
                string dir = folderBrowserDialog1.SelectedPath;
                extractAndInsertData(dir);
            }
        }

        private void extractAndInsertData(string dir)
        {
            var collect = _database.GetCollection<BsonDocument>("docxfiles");
            try
            {
                string[] files = Directory.GetFiles(dir, "*.docx");
                foreach (var fle in files)
                {
                    //open file
                    using (WordprocessingDocument curDoc = WordprocessingDocument.Open(fle, false))
                    {
                        string title;
                        try { title = curDoc.PackageProperties.Title.ToLower(); }
                        catch { title = ""; }

                        string author;
                        try { author = curDoc.PackageProperties.Creator.ToLower(); }
                        catch { author = ""; }

                        string subject;
                        try { subject = curDoc.PackageProperties.Subject.ToLower(); }
                        catch { subject = ""; }

                        string category;
                        try { category = curDoc.PackageProperties.Category.ToLower(); }
                        catch { category = ""; }

                        string keywords;
                        try { keywords = curDoc.PackageProperties.Keywords.ToLower(); }
                        catch { keywords = ""; }

                        string description;
                        try { description = curDoc.PackageProperties.Description.ToLower(); }
                        catch { description = ""; }

                        DateTime? created;
                        try { created = curDoc.PackageProperties.Created; }
                        catch { created = null; }

                        DateTime? modified;
                        try { modified = curDoc.PackageProperties.Modified; }
                        catch { modified = null; }

                        var document = new BsonDocument
                        {
                            {"filename", Path.GetFileNameWithoutExtension(fle)},
                            {"filepath", fle},
                            {"filetype", "docx" },
                            {"title", title},
                            {"author", author},
                            {"subject", subject},
                            {"category", category},
                            {"keywords", keywords},
                            {"description", description},
                            {"created", created},
                            {"modified", modified}
                        };
                        collect.InsertOne(document);
                    }
                }
            }
            catch
            {
                MessageBox.Show("exception");
            }
            MessageBox.Show("Done updating database");
        }
    }
}

using KPLN_CoordiantorAI.Common;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;

namespace KPLN_CoordiantorAI.Forms
{
    public partial class ArticleAliasesWindow : Window
    {
        private readonly ObservableCollection<ArticleAliasRow> _rows;

        public ArticleAliasesWindow(string embeddingsFolderPath, string moodleUrlTemplate, string aliasesJson)
        {
            _rows = new ObservableCollection<ArticleAliasRow>();
            AliasesJson = string.IsNullOrWhiteSpace(aliasesJson) ? ArticleAliasSettings.EmptyJson : aliasesJson;

            InitializeComponent();
            ArticlesDataGrid.ItemsSource = _rows;
            LoadRows(embeddingsFolderPath, moodleUrlTemplate, AliasesJson);
        }

        public string AliasesJson { get; private set; }

        private void LoadRows(string embeddingsFolderPath, string moodleUrlTemplate, string aliasesJson)
        {
            IDictionary<string, string> aliasesByArticleId = LoadAliases(aliasesJson);
            try
            {
                IList<MoodleArticleInfo> articles = MoodleEmbeddingIndex.LoadArticleInfos(embeddingsFolderPath, moodleUrlTemplate);
                foreach (MoodleArticleInfo article in articles)
                {
                    if (article == null || string.IsNullOrWhiteSpace(article.ArticleId))
                        continue;

                    string aliases;
                    aliasesByArticleId.TryGetValue(article.ArticleId, out aliases);
                    _rows.Add(new ArticleAliasRow
                    {
                        ArticleId = article.ArticleId,
                        Title = article.Title,
                        Aliases = aliases
                    });
                }

                StatusTextBlock.Text = _rows.Count == 0
                    ? "Статьи не найдены. Проверьте папку с эмбеддингами и файл moodle_meta.json."
                    : "Статей загружено: " + _rows.Count;
            }
            catch (Exception ex)
            {
                StatusTextBlock.Text = "Не удалось загрузить статьи: " + ex.Message;
            }
        }

        private static IDictionary<string, string> LoadAliases(string aliasesJson)
        {
            try
            {
                return ArticleAliasSettings.ToAliasMap(ArticleAliasSettings.FromJson(aliasesJson));
            }
            catch
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }
        }

        private void OnSaveClick(object sender, RoutedEventArgs e)
        {
            List<ArticleAliasEntry> entries = new List<ArticleAliasEntry>();
            foreach (ArticleAliasRow row in _rows)
            {
                if (row == null || string.IsNullOrWhiteSpace(row.ArticleId) || string.IsNullOrWhiteSpace(row.Aliases))
                    continue;

                entries.Add(new ArticleAliasEntry
                {
                    ArticleId = row.ArticleId,
                    Aliases = row.Aliases
                });
            }

            AliasesJson = ArticleAliasSettings.ToJson(entries);
            DialogResult = true;
        }

        private void OnCancelClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        public sealed class ArticleAliasRow
        {
            public string ArticleId { get; set; }

            public string Title { get; set; }

            public string Aliases { get; set; }
        }
    }
}
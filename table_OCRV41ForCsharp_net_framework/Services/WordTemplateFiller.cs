using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using NPOI.XWPF.UserModel;
using table_OCRV41ForCsharp_net_framework.Models;

namespace table_OCRV41ForCsharp_net_framework.Services
{
    public interface IWordTemplateFiller
    {
        void FillTemplate(XWPFDocument document, OcrResult result);
        void FillParagraphs(IList<XWPFParagraph> paragraphs, OcrResult result);
        void FillTables(IList<XWPFTable> tables, OcrResult result);
    }

    public class WordTemplateFiller : IWordTemplateFiller
    {
        private readonly Dictionary<string, PropertyInfo> _propertyCache;

        public WordTemplateFiller()
        {
            _propertyCache = typeof(OcrResult).GetProperties()
                .ToDictionary(p => p.Name, p => p, StringComparer.OrdinalIgnoreCase);
        }

        public void FillTemplate(XWPFDocument document, OcrResult result)
        {
            FillParagraphs(document.Paragraphs, result);
            FillTables(document.Tables, result);
        }

        public void FillParagraphs(IList<XWPFParagraph> paragraphs, OcrResult result)
        {
            foreach (var paragraph in paragraphs)
            {
                string text = paragraph.ParagraphText;
                if (string.IsNullOrEmpty(text)) continue;

                foreach (var mapping in PlaceholderMappings.Mappings)
                {
                    if (text.Contains(mapping.Placeholder))
                    {
                        string value = GetValue(result, mapping);
                        ReplacePlaceholderInParagraph(paragraph, mapping.Placeholder, value);
                    }
                }
            }
        }

        public void FillTables(IList<XWPFTable> tables, OcrResult result)
        {
            foreach (var table in tables)
            {
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.GetTableCells())
                    {
                        foreach (var paragraph in cell.Paragraphs)
                        {
                            string text = paragraph.ParagraphText;
                            if (string.IsNullOrEmpty(text)) continue;

                            foreach (var mapping in PlaceholderMappings.Mappings)
                            {
                                if (text.Contains(mapping.Placeholder))
                                {
                                    string value = GetValue(result, mapping);
                                    ReplacePlaceholderInParagraph(paragraph, mapping.Placeholder, value);
                                }
                            }
                        }
                    }
                }
            }
        }

        private string GetValue(OcrResult result, PlaceholderMapping mapping)
        {
            if (!_propertyCache.TryGetValue(mapping.PropertyName, out var property))
            {
                return string.Empty;
            }

            var rawValue = property.GetValue(result)?.ToString() ?? string.Empty;

            if (mapping.NeedsPrefix)
            {
                bool isJianyan = result.JianyanOrjiance?.Equals("检验") == true;
                string prefix = isJianyan ? mapping.Prefix检验 : mapping.Prefix检测;
                rawValue = prefix + rawValue;
            }

            if (mapping.AppendUnit && !string.IsNullOrEmpty(rawValue))
            {
                rawValue = rawValue + mapping.Unit;
            }

            return rawValue;
        }

        private void ReplacePlaceholderInParagraph(XWPFParagraph paragraph, string placeholder, string value)
        {
            string text = paragraph.ParagraphText;
            if (!text.Contains(placeholder)) return;

            foreach (var run in paragraph.Runs)
            {
                string runText = run.Text;
                if (runText.Contains(placeholder))
                {
                    // 替换占位符
                    runText = runText.Replace(placeholder, value);
                    
                    // 如果是第一页的委托单位或日期，保持下划线长度一致
                    if (placeholder == "[2]" || placeholder == "[3]")
                    {
                        // 计算需要补充的下划线长度
                        int currentDisplayWidth = GetDisplayWidth(value);
                        int targetDisplayWidth = GetDisplayWidth("广东省特种设备检测研究院江门检测院"); // 以校验单位长度为标准
                        
                        if (currentDisplayWidth < targetDisplayWidth)
                        {
                            // 补充下划线
                            int underlineWidth = targetDisplayWidth - currentDisplayWidth;
                            int underlineCount = (underlineWidth + 1) / 2; // 下划线字符数
                            string underline = new string('_', underlineCount);
                            runText += underline;
                        }
                    }
                    
                    run.SetText(runText);
                    return;
                }
            }
        }

        private int GetDisplayWidth(string text)
        {
            int width = 0;
            foreach (char c in text)
            {
                if (IsChinese(c))
                {
                    width += 2; // 中文字符宽度为2
                }
                else
                {
                    width += 1; // 英文字符和数字宽度为1
                }
            }
            return width;
        }

        private bool IsChinese(char c)
        {
            return (c >= 0x4E00 && c <= 0x9FFF) ||  // 基本汉字
                   (c >= 0x3400 && c <= 0x4DBF) ||  // 扩展A
                   (c >= 0x20000 && c <= 0x2A6DF) || // 扩展B
                   (c >= 0x2A700 && c <= 0x2B73F) || // 扩展C
                   (c >= 0x2B740 && c <= 0x2B81F) || // 扩展D
                   (c >= 0x2B820 && c <= 0x2CEAF) || // 扩展E
                   (c >= 0x2CEB0 && c <= 0x2EBEF) || // 扩展F
                   (c >= 0x30000 && c <= 0x313AF);   // 扩展G
        }
    }
}

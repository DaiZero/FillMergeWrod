using Aspose.Words;
using Aspose.Words.Replacing;
using System.IO;

namespace AutoFillWrodDoc
{
    public class AsposeWordHepler
    {
        public void Copy(string sourceFileName, string destFileName)
        {
            File.Copy(sourceFileName, destFileName, true);
        }

        public void MergeDocument(string fp1, string fp2, string fp3)
        {
            FileStream fs1 = new FileStream(fp1, FileMode.Open);
            Document doc1 = new Document(fs1);
            FileStream fs2 = new FileStream(fp2, FileMode.Open);
            Document doc2 = new Document(fs2);
            Document doc3 = new Document();
            doc3.RemoveAllChildren();
            doc3.AppendDocument(doc1, ImportFormatMode.UseDestinationStyles);
            fs1.Close();
            fs2.Close();
            fs1.Dispose();
            fs2.Dispose();
            doc3.AppendDocument(doc2, ImportFormatMode.UseDestinationStyles);
            doc3.Save(fp3, SaveFormat.Docx);

            //注意：默认的word文档的追加方式在新的一页追加，如果需要在本页增加需要设置word；
            //1 以word2013为例，在头部的选项卡中选择"页面布局"->“纸张大小”->"其它页面大小"
            //2 在页面设置对话框中选择"版式"->节的起始位置中"接续本页"即可.
        }


        public void MergeDocument(string headFileName, string addFileName)
        {
            using (FileStream fs1 = new FileStream(headFileName, FileMode.Open))
            {
                Document doc1 = new Document(fs1);
                using (FileStream fs2 = new FileStream(addFileName, FileMode.Open))
                {
                    Document doc2 = new Document(fs2);
                    fs2.Close();
                    doc1.AppendDocument(doc2, ImportFormatMode.UseDestinationStyles);
                }
                fs1.Close();
                doc1.Save(headFileName, SaveFormat.Docx);
            }
        }

        /// <summary>
        /// 替换(如果新的为空或者NULL值则不替换)
        /// https://apireference.aspose.com/net/words/aspose.words/range/methods/replace
        /// </summary>
        /// <param name="filepath">文件路径</param>
        /// <param name="oldstr">需被替换的旧字符串</param>
        /// <param name="newstr">替换的新字符串</param>
        public void ReplaceString(string filepath, string oldstr, string newstr)
        {
            if(string.IsNullOrWhiteSpace(newstr))
            {
                return;
            }
            using (FileStream fs = new FileStream(filepath, FileMode.Open))
            {
                Document doc = new Document(fs);
                doc.Range.Replace(oldstr, newstr, new FindReplaceOptions());
                doc.Save(filepath, SaveFormat.Docx);
            }
        }

        public void ReplaceString(Document doc, string oldstr, string newstr)
        {
            if (string.IsNullOrWhiteSpace(newstr))
            {
                return;
            }
            doc.Range.Replace(oldstr, newstr, new FindReplaceOptions());
        }

        #region 新增并且写入加替换
        //Document doc = new Document(filepath);
        //DocumentBuilder builder = new DocumentBuilder(doc);
        //builder.Writeln("Numbers 1, 2, 3");
        //doc.Save(filepath, SaveFormat.Docx);
        #endregion
    }
}

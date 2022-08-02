using DevExpress.Pdf;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Office.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace track_changes
{

    
    class Program
    {

        static void Main(string[] args)


        {
            
                RichEditDocumentServer documentProcessor = new RichEditDocumentServer();

                documentProcessor.LoadDocument("t2.docx");
                string commentAuthor = "Victoria";
                documentProcessor.TrackedMovesConflict += DocumentProcessor_TrackedMovesConflict;


                DocumentTrackChangesOptions documentTrackChangesOptions = documentProcessor.Document.TrackChanges;
                documentTrackChangesOptions.Enabled = true;
                documentTrackChangesOptions.TrackFormatting = true;
                documentTrackChangesOptions.TrackMoves = true;
           

                // get list

                var paragraps = "Incapacidade técnica, negligência, imprudência ou imperícia grave por parte da Contratada, seus empregados ou eventuais subcontratados, reiterada e devidamente comprovada durante a execução do Objeto";

                //Formata frases especificas no documento
                //Esta modificacao e adicionada em uma nova revisao

                Document document = documentProcessor.Document;

                DocumentRange[] targetPhrases = documentProcessor.Document.FindAll(paragraps, SearchOptions.None);
          
                
                CharacterProperties characterProperties = documentProcessor.Document.BeginUpdateCharacters(targetPhrases[0]);
                
                characterProperties.FontName = "Arial";
                characterProperties.Italic = true;
                characterProperties.BackColor = Color.Red;
                documentProcessor.Document.EndUpdateCharacters(characterProperties);

                TrackChangesOptions trackChangesOptions = documentProcessor.Options.Annotations.TrackChanges;
                DocumentRange[] targetPhrases2 = documentProcessor.Document.FindAll(paragraps, SearchOptions.None);
                document.Delete(targetPhrases[0]);

                document.Replace(targetPhrases2[0], "\r\n" + paragraps + "\r\n" + "Paragrafo Inserido");

                document.Comments.Create(targetPhrases[0], commentAuthor, DateTime.Now);
                
                int commentCount = document.Comments.Count;
                if (commentCount > 0)
                {
                    
                    Comment comment = document.Comments[document.Comments.Count - 1];
                    if (comment != null)
                    {
                        
                        SubDocument commentDocument = comment.BeginUpdate();                        
                        commentDocument.InsertText(commentDocument.CreatePosition(0), "Alterado pelo Robo");                        
                        commentDocument.Tables.Create(commentDocument.CreatePosition(20), 19, 1);                        
                        comment.EndUpdate(commentDocument);
                    }
                }



                documentProcessor.Options.Annotations.Author = "Victoria";

                //Especifica como as revisoes devem aparecer:


                trackChangesOptions.DisplayForReviewMode = DisplayForReviewMode.AllMarkup;
                trackChangesOptions.DisplayFormatting = DisplayFormatting.ColorOnly;
                trackChangesOptions.FormattingColor = RevisionColor.ClassicBlue;
                trackChangesOptions.DisplayInsertionStyle = DisplayInsertionStyle.None;
                trackChangesOptions.InsertionColor = RevisionColor.DarkRed;

                RevisionCollection documentRevisions = documentProcessor.Document.Revisions;

                //Aceita as revisoes
                SubDocument header = documentProcessor.Document.Sections[0].BeginUpdateHeader(HeaderFooterType.First);
                documentRevisions.AcceptAll(header);
                documentProcessor.Document.Sections[0].EndUpdateHeader(header);

               
                //Rejeita as revisoes
                var sectionRevisions = documentRevisions.Get(documentProcessor.Document.Sections[0].Range).Where(x => x.Author == "Victoria");

                foreach (Revision revision in sectionRevisions)
                revision.Reject();

                //Aceita todas as revisoes
                documentRevisions.AcceptAll(x => x.Type == RevisionType.CharacterPropertyChanged || x.Type == RevisionType.ParagraphPropertyChanged || x.Type == RevisionType.SectionPropertyChanged);

                
                documentProcessor.SaveDocument("Contrato 01.docx", DocumentFormat.OpenXml);
                System.Diagnostics.Process.Start("Contrato 01.docx");

            
        }

        private static void DocumentProcessor_TrackedMovesConflict(object sender, TrackedMovesConflictEventArgs e)
        {
            e.ResolveMode = (e.OriginalLocationRange.Length <= e.NewLocationRange.Length) ? TrackedMovesConflictResolveMode.KeepOriginalLocationText : TrackedMovesConflictResolveMode.KeepNewLocationText;
        }


    }

}


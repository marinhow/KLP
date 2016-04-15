using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KLPOutlookAddIn
{
    class MetadataInfo
    {
        public string Ankomstdato { get; set; }
        public string Indekseringsnokkel { get; set; }
        public string Dokumentkode { get; set; }
        public string DokumentkodeBeskrivelse { get; set; }
        public string Dokumentbeskrivelse { get; set; }
        public string Fodselsnr { get; set; }
        public string DokAnkomstStatus { get; set; }
        public string ExternalLink { get; set; }
        public string Folder { get; set; }
        public bool Validated { get; set; }

        public MetadataInfo(string Ankomstdato, string Indekseringsnokkel, string Dokumentkode, string DokumentkodeBeskrivelse,
            string Dokumentbeskrivelse, string Fodselsnr, string DokAnkomstStatus, string ExternalLink, string Folder, bool validated)
        {
            this.Ankomstdato = Ankomstdato;
            this.Indekseringsnokkel = Indekseringsnokkel;
            this.Dokumentkode = Dokumentkode;
            this.DokumentkodeBeskrivelse = DokumentkodeBeskrivelse;
            this.Dokumentbeskrivelse = Dokumentbeskrivelse;
            this.Fodselsnr = Fodselsnr;
            this.DokAnkomstStatus = DokAnkomstStatus;
            this.ExternalLink = ExternalLink;
            this.Folder = Folder;
            this.Validated = validated;
        }
    }
}

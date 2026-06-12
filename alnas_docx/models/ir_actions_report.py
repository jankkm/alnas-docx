import base64
import zipfile
import os
import re
import subprocess
import tempfile
import shutil
from urllib.parse import urlparse
from io import BytesIO
from functools import partial
import requests
from docx import Document
from docxtpl import DocxTemplate
from docxcompose.composer import Composer

from odoo import _, api, fields, models
from odoo.tools.safe_eval import safe_eval, time
from odoo.exceptions import ValidationError, MissingError, UserError

from ..tools import misc as misc_tools


class IrActionsReport(models.Model):
    _inherit = "ir.actions.report"

    report_type = fields.Selection(
        selection_add=[("docx", "DOCX")], ondelete={"docx": "cascade"}
    ) # add docx type
    report_docx_template = fields.Binary(string="Report DOCX Template")
    report_docx_template_name = fields.Char(string="Report DOCX Template Name")
    docx_merge_mode = fields.Selection(
        [
            ("composer", "Composer"),
            ("zip", "Zip"),
            ("pdf", "PDF"),
            ("pdf_in_zip", "PDF in Zip"),
        ],
        string="DOCX Mode",
        default="composer",
    )
    docx_autoescape = fields.Boolean(
        string="Autoescape (Docx)",
        default=False,
        help="Enable autoescape for special character like <, > and &.",
    )

    @api.constrains("report_type")
    def _check_report_type(self):
        for rec in self:
            if (
                rec.report_type == "docx"
                and not rec.report_docx_template
                and not rec.report_docx_template_name.endswith(".docx")
            ):
                raise ValidationError(_("Please upload a DOCX template."))

    def _get_rendering_context_docx(self, doc_template, extra_pdfs=None):
        context = {
            "company": self.env.company,
            "lang": self.env.context.get("lang", "id_ID"),
            "sysdate": fields.Datetime.now(),
            "spelled_out": misc_tools.spelled_out,
            "parsehtml": misc_tools.parse_html,
            "formatdate": misc_tools.formatdate,
            "format_datetime": misc_tools.format_datetime,
            "convert_currency": misc_tools.convert_currency,
            "formatabs": misc_tools.format_abs,
            "rich_text": misc_tools.rich_text,
            "render_image": partial(misc_tools.render_image, doc_template),
            "html2docx": partial(misc_tools.render_html_as_subdoc, doc_template),
            "add_subdoc": partial(misc_tools.add_new_subdoc, doc_template),
            "replace_image": partial(misc_tools.replace_image, doc_template),
            "replace_media": partial(misc_tools.replace_media, doc_template),
            "replace_embedded": partial(misc_tools.replace_embedded, doc_template),
            "replace_zipname": partial(misc_tools.replace_zipname, doc_template),
            "linked_attachments": lambda record: misc_tools.linked_attachments_for_record(
                self.env, record
            ),
            "add_pdf": (
                misc_tools.add_pdf_factory(extra_pdfs["before"], extra_pdfs["after"])
                if extra_pdfs is not None
                else (lambda *args, **kwargs: "")
            ),
        }
        return context
    
    def _render_docx(self, report_ref, docids, data):
        report = self._get_report(report_ref)
        template = report.report_docx_template

        if not template:
            raise MissingError("No DOCX template found.")

        doc_template = DocxTemplate(BytesIO(base64.b64decode(template)))
        doc_obj = self.env[report.model].browse(docids).with_context(
            bin_size=False
        )
        extra_pdfs = (
            {"before": [], "after": []}
            if report.docx_merge_mode in ("pdf", "pdf_in_zip")
            else None
        )
        context = self._get_rendering_context_docx(
            doc_template=doc_template, extra_pdfs=extra_pdfs
        )
        autoescape = report.docx_autoescape
        
        if report.docx_merge_mode == "composer":
            return self._render_composer_mode(doc_template, doc_obj, data, context, autoescape=autoescape)
        elif report.docx_merge_mode == "zip":
            return self._render_zip_mode(
                doc_template,
                doc_obj,
                data,
                context,
                report_name=report.print_report_name,
                autoescape=autoescape,
            )
        elif report.docx_merge_mode == "pdf_in_zip":
            return self._render_pdf_in_zip_mode(
                doc_template,
                doc_obj,
                data,
                context,
                extra_pdfs,
                report_name=report.print_report_name,
                autoescape=autoescape,
            )
        else:
            return self._render_docx_to_pdf_mode(
                doc_template, doc_obj, data, context, extra_pdfs, autoescape=autoescape
            )

    def _render_composer_mode(self, doc_template, doc_obj, data, context, autoescape=False):
        for idx, obj in enumerate(doc_obj):
            context = {
                **context,
                "docs": obj,
                "data": data
            }

            temp = BytesIO()
            doc_template.render(context, autoescape=autoescape)
            doc_template.save(temp)
            temp.seek(0)

            if len(doc_obj) == 1:
                return temp.read(), 'docx'
            else:
                if idx == 0:
                    master_doc = Document(temp)
                    composer = Composer(master_doc)
                else:
                    doc_to_append = Document(temp)
                    master_doc.add_page_break()
                    composer.append(doc_to_append)

        temp_output = BytesIO()
        composer.save(temp_output)
        temp_output.seek(0)

        return temp_output.read(), 'docx'

    def _render_zip_mode(
        self, doc_template, doc_obj, data, context, report_name="report", autoescape=False
    ):
        docx_files = []

        for obj in doc_obj:
            context = {
                **context,
                "docs": obj,
                "data": data
            }

            temp = BytesIO()
            doc_template.render(context, autoescape=autoescape)
            doc_template.save(temp)
            temp.seek(0)
            docx_files.append(temp.read())

        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for idx, docx_file in enumerate(docx_files):
                name = safe_eval(report_name, {"object": doc_obj[idx], "time": time})
                filename = f"{self._sanitize_report_filename(name)}.docx"
                zip_file.writestr(filename, docx_file)

        zip_buffer.seek(0)

        return zip_buffer.read(), 'zip'

    def _docx_bytes_to_pdf(self, docx_bytes, extra_pdfs):
        temp_dir = tempfile.mkdtemp()
        try:
            docx_file_path = os.path.join(temp_dir, "document.docx")
            with open(docx_file_path, "wb") as f:
                f.write(docx_bytes)
            pdf_file_path = self.convert_file_to_pdf(docx_file_path, temp_dir)
            if not pdf_file_path:
                raise UserError('PDF conversion failed.')

            with open(pdf_file_path, 'rb') as pdf_file:
                main_pdf = pdf_file.read()

        finally:
            shutil.rmtree(temp_dir)

        if extra_pdfs and (extra_pdfs["before"] or extra_pdfs["after"]):
            main_pdf = misc_tools.merge_pdf_bytes(
                main_pdf, extra_pdfs["before"], extra_pdfs["after"]
            )

        return main_pdf, 'pdf'

    def _render_pdf_in_zip_mode(
        self, doc_template, doc_obj, data, context, extra_pdfs, report_name="report", autoescape=False
    ):
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for idx, obj in enumerate(doc_obj):
                one = doc_obj.filtered(lambda r, oid=obj.id: r.id == oid)
                # Keep add_pdf() side effects isolated per record in ZIP mode.
                record_extra_pdfs = {"before": [], "after": []}
                record_context = self._get_rendering_context_docx(
                    doc_template=doc_template, extra_pdfs=record_extra_pdfs
                )
                pdf_bytes, _ = self._render_docx_to_pdf_mode(
                    doc_template,
                    one,
                    data,
                    record_context,
                    record_extra_pdfs,
                    autoescape=autoescape,
                )
                name = safe_eval(report_name, {"object": doc_obj[idx], "time": time})
                filename = f"{self._sanitize_report_filename(name)}.pdf"
                zip_file.writestr(filename, pdf_bytes)

        zip_buffer.seek(0)
        return zip_buffer.read(), "zip"

    def _render_docx_to_pdf_mode(self, doc_template, doc_obj, data, context, extra_pdfs, autoescape=False):
        docx_file, _ = self._render_composer_mode(
            doc_template, doc_obj, data, context, autoescape=autoescape
        )
        return self._docx_bytes_to_pdf(docx_file, extra_pdfs)

    def convert_file_to_pdf(self, file_path, output_dir):
        uno_url = self._get_libreoffice_uno_url()
        if uno_url:
            self._convert_file_to_pdf_uno(file_path, output_dir, uno_url)
        else:
            librepath = self._get_libreoffice_path()
            subprocess.run(
                [librepath, '--headless', '--convert-to', 'pdf', '--outdir', output_dir, file_path],
                check=True,
            )
        pdf_file_name = os.path.splitext(os.path.basename(file_path))[0] + '.pdf' 
        pdf_file_path = os.path.join(output_dir, pdf_file_name)        
        if os.path.exists(pdf_file_path):
            return pdf_file_path
        else:
            return None

    def _convert_file_to_pdf_uno(self, file_path, output_dir, uno_url):
        base_url = uno_url.strip().rstrip("/")
        if "://" not in base_url:
            base_url = "http://%s" % base_url
        parsed = urlparse(base_url)
        endpoint = base_url if parsed.path and parsed.path != "/" else "%s/request" % base_url
        output_file_path = os.path.join(
            output_dir,
            "%s.pdf" % os.path.splitext(os.path.basename(file_path))[0],
        )
        try:
            with open(file_path, "rb") as input_file:
                response = requests.post(
                    endpoint,
                    files={"file": (os.path.basename(file_path), input_file)},
                    data={"convert-to": "pdf"},
                    timeout=120,
                )
            response.raise_for_status()
        except requests.RequestException as exc:
            raise UserError(
                _(
                    "Failed to convert via UNO URL '%(url)s'. "
                    "Check that the URL is reachable from the Odoo runtime."
                )
                % {"url": endpoint}
            ) from exc
        with open(output_file_path, "wb") as output_file:
            output_file.write(response.content)

    def _sanitize_report_filename(self, name):
        # Normalize any incoming value to a basic filename candidate.
        s = str(name or "").strip()
        # Remove ASCII control chars that can break archive extraction/tools.
        s = re.sub(r"[\x00-\x1f\x7f]", "_", s)
        # Replace cross-platform invalid filename characters.
        s = re.sub(r'[<>:"/\\|?*]', "_", s)
        # Avoid troublesome leading/trailing spaces and dots on Windows.
        s = s.strip(" .")
        if not s:
            # Ensure we always produce a non-empty filename.
            s = "report"

        # Reserved DOS device names cannot be used as filenames.
        reserved = {
            "CON",
            "PRN",
            "AUX",
            "NUL",
            *(f"COM{i}" for i in range(1, 10)),
            *(f"LPT{i}" for i in range(1, 10)),
        }
        if s.upper() in reserved:
            s = f"_{s}"

        # Keep names reasonably short for broad ZIP/tool compatibility.
        return s[:150]

    def _get_libreoffice_path(self):
        libreoffice = self.env["ir.config_parameter"].sudo().get_param("libreoffice.path")
        if not libreoffice:
            raise ValidationError('Libreoffice path doesnt exits, \n \
                please set in Settings => Technical => Parameters => System Parameters => default_libreoffice_path')
            
        return libreoffice

    def _get_libreoffice_uno_url(self):
        return self.env["ir.config_parameter"].sudo().get_param("libreoffice.uno.url", "").strip()

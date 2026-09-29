"""
ansam_doc.py – Génère des documents Word (.docx) au branding ANSAM.

Principe : on part TOUJOURS du template officiel (assets/ANSAM_Modele.docx).
En-têtes, pieds de page, logos, polices (Montserrat) et styles sont conservés ;
seul le corps du document est remplacé par le contenu fourni.

Usage minimal :

    import sys; sys.path.insert(0, "<chemin>/ansam-branding/scripts")
    from ansam_doc import AnsamDoc

    d = AnsamDoc()
    d.cover("Titre du document", "Sous-titre",
            recipient=["Client SA", "Rue X 1", "1000 Lausanne"],
            meta=[("Objet", "…"), ("Version / date", "1.0 – 01.01.2026")])
    d.h1("Introduction", page_break=True)
    d.p("Texte avec **gras**.")
    d.bullets(["Point 1", "Point 2"])
    d.steps(["Étape 1", "Étape 2"])
    d.table(["Col A", "Col B"], [["a", "b"]], widths=[3000, 6072])
    d.info("Titre", ["ligne"]); d.warn("Attention", ["ligne"])
    d.placeholder("[Capture à insérer : …]")
    d.image("capture.png", "Légende")
    d.qr_box()
    d.save("sortie.docx")

Aucune dépendance externe (bibliothèque standard Python 3.8+).
"""
import os
import re
import shutil
import struct
import tempfile
import zipfile
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE = os.path.join(HERE, "..", "assets", "ANSAM_Modele.docx")

# ---- Charte ANSAM ----
PINK = "D7007F"          # rose ANSAM (titres, en-têtes de tableau, accents)
LIGHT_PINK = "FBE5F2"    # fond des encadrés d'information
WARN_FILL = "FFF2CC"     # fond des encadrés d'avertissement
WARN_BAR = "BF8F00"
GREY_FILL = "F2F2F2"     # fond des emplacements à compléter
BORDER = "BFBFBF"        # bordures de tableaux
CAPTION = "595959"
CONTENT_W = 9072         # largeur utile A4 (11906 - 2 x 1417 twips)

BULLET_NUM_ID = 14       # liste à puces du template
DECIMAL_ABSTRACT = 3     # définition « 1. 2. 3. » du template

TITLE_RPR = ('<w:rFonts w:ascii="Montserrat Black" w:eastAsiaTheme="majorEastAsia" w:hAnsi="Montserrat Black" '
             'w:cstheme="majorBidi"/><w:b/><w:color w:val="D7007F"/><w:spacing w:val="-10"/><w:kern w:val="28"/>'
             '<w:sz w:val="56"/><w:szCs w:val="52"/>')
SP = ('<w:spacing w:before="0" w:beforeAutospacing="0" w:after="120" w:afterAutospacing="0" '
      'w:line="278" w:lineRule="auto"/>')
SP0 = '<w:spacing w:before="0" w:beforeAutospacing="0" w:after="0" w:afterAutospacing="0"/>'


def _runs(text, base_rpr=""):
    """Texte -> runs Word. **gras** et `code` sont reconnus."""
    out = []
    for part in re.split(r"(\*\*[^*]+\*\*|`[^`]+`)", text):
        if not part:
            continue
        if part.startswith("**"):
            out.append(f'<w:r><w:rPr>{base_rpr}<w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">{escape(part[2:-2])}</w:t></w:r>')
        elif part.startswith("`"):
            rpr = base_rpr + '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:color w:val="833C0B"/><w:sz w:val="20"/>'
            out.append(f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{escape(part[1:-1])}</w:t></w:r>')
        else:
            out.append(f'<w:r><w:rPr>{base_rpr}</w:rPr><w:t xml:space="preserve">{escape(part)}</w:t></w:r>')
    return "".join(out)


def _borders(style="single", sz=6, color=BORDER):
    return "".join(f'<w:{s} w:val="{style}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
                   for s in ("top", "left", "bottom", "right"))


def _cell(xml, width, fill=None, borders=None, valign=None):
    b = f"<w:tcBorders>{borders}</w:tcBorders>" if borders else ""
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else ""
    va = f'<w:vAlign w:val="{valign}"/>' if valign else ""
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{b}{shd}{va}</w:tcPr>{xml}</w:tc>'


def _img_size(data):
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        w, h = struct.unpack(">II", data[16:24])
        return w, h, "png"
    if data[:2] == b"\xff\xd8":
        i = 2
        while i < len(data):
            if data[i] != 0xFF:
                i += 1
                continue
            m = data[i + 1]
            ln = struct.unpack(">H", data[i + 2:i + 4])[0]
            if 0xC0 <= m <= 0xCF and m not in (0xC4, 0xC8, 0xCC):
                h, w = struct.unpack(">HH", data[i + 5:i + 9])
                return w, h, "jpeg"
            i += 2 + ln
    raise ValueError("Format d'image non supporté (PNG ou JPEG uniquement)")


class AnsamDoc:
    def __init__(self, template=TEMPLATE):
        self.tmp = tempfile.mkdtemp(prefix="ansam_")
        with zipfile.ZipFile(template) as z:
            z.extractall(self.tmp)
        doc = self._read("word/document.xml")
        self.head = doc[: doc.find("<w:body>") + len("<w:body>")]
        body = doc[doc.find("<w:body>") + len("<w:body>"):]
        paras = re.findall(r"<w:p[ >].*?</w:p>", body, re.S)
        self.logo_para = next(p for p in paras if "<w:drawing" in p)  # logo ANSAM group de la couverture
        self.sect = re.search(r"<w:sectPr.*?</w:sectPr>", body, re.S).group(0)
        self.numbering = self._read("word/numbering.xml")
        self.rels = self._read("word/_rels/document.xml.rels")
        self.ctypes = self._read("[Content_Types].xml")
        self.parts = []
        self._next_num = 500
        self._next_img = 900
        self.title = ""

    # ---------- utilitaires fichiers ----------
    def _read(self, rel):
        with open(os.path.join(self.tmp, rel), encoding="utf8") as f:
            return f.read()

    def _write(self, rel, content):
        with open(os.path.join(self.tmp, rel), "w", encoding="utf8") as f:
            f.write(content)

    def raw(self, xml):
        """Ajoute du WordprocessingML brut (usage avancé)."""
        self.parts.append(xml)

    # ---------- blocs ----------
    def empty(self, n=1):
        for _ in range(n):
            self.parts.append(f"<w:p><w:pPr>{SP0}</w:pPr></w:p>")

    def p(self, text, align=None, rpr="", spacing=SP, keep_next=False):
        jc = f'<w:jc w:val="{align}"/>' if align else ""
        kn = "<w:keepNext/>" if keep_next else ""
        self.parts.append(f"<w:p><w:pPr>{kn}{spacing}{jc}</w:pPr>{_runs(text, rpr)}</w:p>")

    def small(self, text, align="left"):
        self.p(text, align, f'<w:i/><w:color w:val="7F7F7F"/><w:sz w:val="16"/>')

    def page_break(self):
        self.parts.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')

    def _heading(self, style, text, page_break):
        br = '<w:r><w:br w:type="page"/></w:r>' if page_break else ""
        self.parts.append(f'<w:p><w:pPr><w:pStyle w:val="{style}"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/>'
                          f'</w:numPr></w:pPr>{br}<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r></w:p>')

    def h1(self, text, page_break=False):
        """Titre de section (style Titre1 du template : Montserrat Black 14 pt, rose)."""
        self._heading("Titre1", text, page_break)

    def h2(self, text, page_break=False):
        """Sous-titre (style Titre2 : gras rose)."""
        self._heading("Titre2", text, page_break)

    def h3(self, text):
        self._heading("Titre3", text, False)

    def bullets(self, items):
        for t in items:
            self.parts.append(f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{BULLET_NUM_ID}"/></w:numPr>'
                              f'<w:spacing w:before="0" w:beforeAutospacing="0" w:after="80" w:afterAutospacing="0" '
                              f'w:line="278" w:lineRule="auto"/><w:jc w:val="left"/></w:pPr>{_runs(t)}</w:p>')

    def steps(self, items):
        """Liste numérotée 1. 2. 3. (redémarre à 1 à chaque appel)."""
        nid = self._next_num
        self._next_num += 1
        i = self.numbering.rfind("</w:num>") + len("</w:num>")
        self.numbering = (self.numbering[:i] + f'<w:num w:numId="{nid}"><w:abstractNumId w:val="{DECIMAL_ABSTRACT}"/>'
                          f'<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>' + self.numbering[i:])
        for t in items:
            self.parts.append(f'<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{nid}"/></w:numPr>'
                              f'<w:spacing w:before="0" w:beforeAutospacing="0" w:after="80" w:afterAutospacing="0" '
                              f'w:line="278" w:lineRule="auto"/><w:jc w:val="left"/></w:pPr>{_runs(t)}</w:p>')

    def table(self, header, rows, widths=None):
        """Tableau : en-tête rose ANSAM / texte blanc gras, bordures grises. widths en twips (somme = 9072)."""
        n = len(header)
        widths = widths or [CONTENT_W // n] * (n - 1) + [CONTENT_W - (CONTENT_W // n) * (n - 1)]
        tw = sum(widths)
        grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
        inside = (f'<w:insideH w:val="single" w:sz="6" w:space="0" w:color="{BORDER}"/>'
                  f'<w:insideV w:val="single" w:sz="6" w:space="0" w:color="{BORDER}"/>')

        def cp(text, head):
            rpr = '<w:color w:val="FFFFFF"/><w:b/><w:bCs/>' if head else ""
            return (f'<w:p><w:pPr><w:spacing w:before="40" w:beforeAutospacing="0" w:after="40" w:afterAutospacing="0"/>'
                    f'<w:jc w:val="left"/></w:pPr>{_runs(str(text), rpr)}</w:p>')

        x = (f'<w:tbl><w:tblPr><w:tblW w:w="{tw}" w:type="dxa"/><w:tblBorders>{_borders()}{inside}</w:tblBorders>'
             f'<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/>'
             f'</w:tblCellMar></w:tblPr><w:tblGrid>{grid}</w:tblGrid>')
        x += "<w:tr><w:trPr><w:tblHeader/></w:trPr>" + "".join(_cell(cp(h, True), w, PINK) for h, w in zip(header, widths)) + "</w:tr>"
        for r in rows:
            x += "<w:tr><w:trPr><w:cantSplit/></w:trPr>" + "".join(_cell(cp(c, False), w) for c, w in zip(r, widths)) + "</w:tr>"
        self.parts.append(x + "</w:tbl>")
        self.empty()

    def _box(self, title, lines, fill, bar):
        t_rpr = f'<w:b/><w:bCs/><w:color w:val="{bar}"/>'
        c = (f'<w:p><w:pPr><w:spacing w:before="60" w:beforeAutospacing="0" w:after="80" w:afterAutospacing="0"/>'
             f'<w:jc w:val="left"/></w:pPr>{_runs(title, t_rpr)}</w:p>')
        c += "".join(f'<w:p><w:pPr><w:spacing w:before="0" w:beforeAutospacing="0" w:after="60" w:afterAutospacing="0"/>'
                     f'<w:jc w:val="left"/></w:pPr>{_runs(l)}</w:p>' for l in lines)
        b = (f'<w:top w:val="nil"/><w:left w:val="single" w:sz="36" w:space="0" w:color="{bar}"/>'
             f'<w:bottom w:val="nil"/><w:right w:val="nil"/>')
        self.parts.append(f'<w:tbl><w:tblPr><w:tblW w:w="{CONTENT_W}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
                          f'<w:tblCellMar><w:left w:w="200" w:type="dxa"/><w:right w:w="200" w:type="dxa"/></w:tblCellMar>'
                          f'</w:tblPr><w:tblGrid><w:gridCol w:w="{CONTENT_W}"/></w:tblGrid><w:tr><w:trPr><w:cantSplit/>'
                          f'</w:trPr>{_cell(c, CONTENT_W, fill, b)}</w:tr></w:tbl>')
        self.empty()

    def info(self, title, lines):
        """Encadré d'information : fond rose clair, barre rose à gauche."""
        self._box(title, lines, LIGHT_PINK, PINK)

    def warn(self, title, lines):
        """Encadré d'avertissement : fond jaune, barre ocre."""
        self._box("⚠ " + title, lines, WARN_FILL, WARN_BAR)

    def placeholder(self, label, height=2200):
        """Cadre gris pointillé « à compléter » (capture, photo…)."""
        rpr = f'<w:i/><w:color w:val="{CAPTION}"/><w:sz w:val="20"/>'
        c = f'<w:p><w:pPr>{SP0}<w:jc w:val="center"/></w:pPr>{_runs(label, rpr)}</w:p>'
        self.parts.append(f'<w:tbl><w:tblPr><w:tblW w:w="{CONTENT_W}" w:type="dxa"/><w:jc w:val="center"/>'
                          f'<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="{CONTENT_W}"/></w:tblGrid>'
                          f'<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="{height}" w:hRule="atLeast"/></w:trPr>'
                          f'{_cell(c, CONTENT_W, GREY_FILL, _borders("dashed", 12, "7F7F7F"), "center")}</w:tr></w:tbl>')
        self.empty()

    def qr_box(self, subtitle="(QR code d’enrôlement)"):
        """Cadre carré 8 cm à double bordure rose : « COLLER ICI le QR code »."""
        w = 4536
        l1 = '<w:rFonts w:ascii="Montserrat Black" w:hAnsi="Montserrat Black"/><w:b/><w:color w:val="D7007F"/><w:sz w:val="36"/>'
        l2 = '<w:color w:val="D7007F"/>'
        l3 = f'<w:i/><w:color w:val="{CAPTION}"/><w:sz w:val="18"/>'
        c = (f'<w:p><w:pPr>{SP0}<w:jc w:val="center"/></w:pPr>{_runs("COLLER ICI", l1)}</w:p>'
             f'<w:p><w:pPr>{SP0}<w:jc w:val="center"/></w:pPr>{_runs("le QR code", l2)}</w:p>'
             f'<w:p><w:pPr>{SP0}<w:jc w:val="center"/></w:pPr>{_runs(subtitle, l3)}</w:p>')
        self.parts.append(f'<w:tbl><w:tblPr><w:tblW w:w="{w}" w:type="dxa"/><w:jc w:val="center"/><w:tblLayout w:type="fixed"/>'
                          f'</w:tblPr><w:tblGrid><w:gridCol w:w="{w}"/></w:tblGrid><w:tr><w:trPr><w:cantSplit/>'
                          f'<w:trHeight w:val="{w}" w:hRule="exact"/></w:trPr>'
                          f'{_cell(c, w, None, _borders("double", 18, PINK), "center")}</w:tr></w:tbl>')
        self.empty()

    def image(self, path, caption="", width_cm=15.0, max_height_cm=11.0, border=True):
        """Image centrée (PNG/JPEG) + légende en italique gris."""
        data = open(path, "rb").read()
        w, h, ext = _img_size(data)
        self._next_img += 1
        n = self._next_img
        rid, name = f"rIdAnsamImg{n}", f"ansam_img_{n}.{ext}"
        with open(os.path.join(self.tmp, "word", "media", name), "wb") as f:
            f.write(data)
        self.rels = self.rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/'
                                      f'officeDocument/2006/relationships/image" Target="media/{name}"/></Relationships>')
        if f'Extension="{ext}"' not in self.ctypes:
            self.ctypes = self.ctypes.replace("<Default ", f'<Default Extension="{ext}" ContentType="image/{ext}"/><Default ', 1)
        emu = 360000
        cx, cy = width_cm * emu, width_cm * emu * h / w
        if cy > max_height_cm * emu:
            cy = max_height_cm * emu
            cx = cy * w / h
        cx, cy = int(cx), int(cy)
        ln = f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{BORDER}"/></a:solidFill></a:ln>' if border else ""
        desc = escape(caption, {'"': "&quot;"})
        drawing = (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
                   f'<wp:docPr id="{n}" name="Image {n}" descr="{desc}"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/'
                   f'drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                   f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr>'
                   f'<pic:cNvPr id="{n}" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{rid}"/>'
                   f'<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/>'
                   f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{ln}</pic:spPr>'
                   f'</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')
        self.parts.append(f'<w:p><w:pPr><w:keepNext/><w:spacing w:before="120" w:beforeAutospacing="0" w:after="40" '
                          f'w:afterAutospacing="0"/><w:jc w:val="center"/></w:pPr>{drawing}</w:p>')
        if caption:
            self.p(caption, "center", f'<w:i/><w:color w:val="{CAPTION}"/><w:sz w:val="18"/>')

    # ---------- couverture ----------
    def cover(self, title, subtitle="", recipient=None, meta=None):
        """Page de garde ANSAM : logo du groupe, titre rose, « À L’ATTENTION DE », tableau récapitulatif."""
        self.title = title
        self.parts.append(self.logo_para)
        self.empty(2)
        self.p(title, "center", TITLE_RPR)
        if subtitle:
            self.p(subtitle, "center", '<w:sz w:val="26"/>')
        if recipient:
            self.empty(2)
            self.p("**À L’ATTENTION DE :**", "center", spacing=SP0)
            for line in recipient:
                self.p(f"**{line}**", "center", spacing=SP0)
        if meta:
            self.empty(2)
            self.table(["Document", "Détail"], meta, [2800, CONTENT_W - 2800])

    # ---------- sauvegarde ----------
    def save(self, out_path):
        self._write("word/document.xml", self.head + "".join(self.parts) + self.sect + "</w:body></w:document>")
        self._write("word/numbering.xml", self.numbering)
        self._write("word/_rels/document.xml.rels", self.rels)
        self._write("[Content_Types].xml", self.ctypes)
        if self.title:
            core = self._read("docProps/core.xml")
            if "<dc:title>" in core:
                core = re.sub(r"<dc:title>.*?</dc:title>", f"<dc:title>{escape(self.title)}</dc:title>", core, flags=re.S)
            else:
                core = core.replace("</cp:coreProperties>", f"<dc:title>{escape(self.title)}</dc:title></cp:coreProperties>")
            self._write("docProps/core.xml", core)
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        if os.path.exists(out_path):
            os.remove(out_path)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(os.path.join(self.tmp, "[Content_Types].xml"), "[Content_Types].xml")
            for root, _, files in os.walk(self.tmp):
                for f in files:
                    full = os.path.join(root, f)
                    arc = os.path.relpath(full, self.tmp).replace(os.sep, "/")
                    if arc != "[Content_Types].xml":
                        z.write(full, arc)
        shutil.rmtree(self.tmp, ignore_errors=True)
        return out_path

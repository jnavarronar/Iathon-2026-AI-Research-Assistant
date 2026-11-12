import os
import re
import json
import argparse
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
from openai import OpenAI

def delete_slide(presentation, index):
    """Deletes a slide from the presentation by its index."""
    xml_slides = presentation.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[index])

def extract_title_image(md_content):
    """
    Intenta detectar la imagen de portada (la que insertaste debajo del H1).
    Devuelve el path relativo (ej: 'figures/title_image.jpg') o None.
    """
    # Caso 1: exactamente tu nombre esperado
    m = re.search(r'!\[\]\((figures/title_image\.(png|jpg|jpeg))\)', md_content, re.IGNORECASE)
    if m:
        return m.group(1)

    # Caso 2: si algún día no se llama title_image, toma la primera imagen del documento
    imgs = re.findall(r'!\[.*?\]\((.*?)\)', md_content)
    return imgs[0] if imgs else None

def parse_markdown(md_content):
    """
    Parsea el Markdown en SECCIONES por H2 (## ...).
    Devuelve:
      - sections: lista de dicts {title, level, text, images}
      - all_images: lista de todas las imágenes encontradas
      - title_h1: título principal si existe
    """
    # Todas las imágenes (para contexto global)
    all_images = re.findall(r'!\[.*?\]\((.*?)\)', md_content)

    # Extraer H1 si existe
    h1 = None
    m = re.search(r'^\#\s+(.+)$', md_content, flags=re.MULTILINE)
    if m:
        h1 = m.group(1).strip()

    # Dividir por H2
    # Captura: "## Titulo" y el contenido hasta el siguiente "## "
    pattern = r'(^##\s+.+$)(.*?)(?=^##\s+|\Z)'
    matches = re.findall(pattern, md_content, flags=re.MULTILINE | re.DOTALL)

    sections = []

    # Si no hay H2, hacemos una sección única con todo el texto
    if not matches:
        sections.append({
            "title": h1 or "Report",
            "level": 2,
            "text": _clean_md_text(md_content),
            "images": all_images
        })
        return sections, all_images, h1

    for header, body in matches:
        title = header.replace("##", "").strip()

        # imágenes dentro de esta sección
        imgs = re.findall(r'!\[.*?\]\((.*?)\)', body)

        # texto limpio (sin imágenes)
        clean_text = _clean_md_text(body)

        sections.append({
            "title": title,
            "level": 2,
            "text": clean_text.strip(),
            "images": imgs
        })

    return sections, all_images, h1

def _clean_md_text(text):
    """Quita imágenes y reduce ruido para que el LLM no se vaya solo a figuras."""
    # elimina bloques ::: ... ::: (como tu portada)
    text = re.sub(r':::\s*\{.*?\}\s*.*?:::\s*', '', text, flags=re.DOTALL)

    # elimina tags de imagen markdown
    text = re.sub(r'!\[.*?\]\(.*?\)', '', text)

    # elimina líneas vacías repetidas
    text = re.sub(r'\n{3,}', '\n\n', text)

    return text

def estimate_text_capacity(placeholder):
    """Estimates how many bullet points and characters a placeholder can hold."""
    width_inches = placeholder.width / 914400
    height_inches = placeholder.height / 914400
    
    # More aggressive estimation - typical slides can hold more text
    # Assume ~10pt font, ~8-10 lines per inch vertically
    # Each bullet takes about 0.3-0.35 inches with spacing
    max_bullets = max(3, int(height_inches / 0.3))
    
    # Estimate max chars per bullet based on width
    # At 10-12pt, you can fit roughly 15-18 chars per inch
    max_chars_per_bullet = max(50, int(width_inches * 15))
    
    # Cap at reasonable maximums to maintain readability
    max_bullets = min(max_bullets, 10)
    max_chars_per_bullet = min(max_chars_per_bullet, 120)
    
    return max_bullets, max_chars_per_bullet

def get_template_layouts(pptx_path):
    """Extracts detailed metadata from the template to help the LLM select layouts."""
    prs = Presentation(pptx_path)
    layouts = []
    for i, layout in enumerate(prs.slide_layouts):
        placeholders = []
        text_capacity = None
        
        for p in layout.placeholders:
            ph_info = {
                "name": p.name,
                "type": str(p.placeholder_format.type),
                "left": p.left,
                "top": p.top,
                "width": p.width,
                "height": p.height
            }
            placeholders.append(ph_info)
            
            if p.placeholder_format.type in [PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT]:
                max_bullets, max_chars = estimate_text_capacity(p)
                text_capacity = {
                    "max_bullets": max_bullets,
                    "max_chars_per_bullet": max_chars
                }
        
        layouts.append({
            "index": i,
            "name": layout.name,
            "placeholders": placeholders,
            "placeholder_count": len(placeholders),
            "text_capacity": text_capacity
        })
    return layouts

def auto_fit_text(text_frame, max_font_size=18, min_font_size=10):
    """Automatically adjusts font size to fit content in text frame."""
    if not text_frame.text.strip():
        return
    
    for size in range(max_font_size, min_font_size - 1, -1):
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(size)
        
        if text_frame.text:
            break

def summarize_and_map_with_llm(sections, layouts, image_list, api_key, title_h1=None):
    """Consults the LLM to structure the report and pick the best template layouts."""
    client = OpenAI(api_key=api_key)
    
    layout_descriptions = []
    for l in layouts:
        ph_types = [p['type'] for p in l['placeholders']]
        has_picture = any('PICTURE' in t for t in ph_types)
        has_body = any('BODY' in t or 'OBJECT' in t for t in ph_types)
        
        desc = f"- Index {l['index']}: '{l['name']}'"
        
        if l.get('text_capacity'):
            cap = l['text_capacity']
            desc += f" [Max: {cap['max_bullets']} bullets, {cap['max_chars_per_bullet']} chars/bullet]"
        
        if has_picture:
            desc += " [IMAGES]"
        if has_body:
            desc += " [TEXT]"
            
        layout_descriptions.append(desc)
    
    layout_context = "\n".join(layout_descriptions)

    # Convertimos secciones a JSON corto para el LLM
    sections_payload = []
    for s in sections:
        sections_payload.append({
            "section_title": s["title"],
            "section_text": s["text"][:4000],
            "section_images": s["images"][:8]
        })

    sections_json = json.dumps(sections_payload, ensure_ascii=False, indent=2)
    deck_title = title_h1 or "Report"

    prompt = f"""
    You are a professional presentation designer.

    HARD RULES:
    1) Slide 1 must be kind="title" (title: deck title). No bullets, no images.
    2) Slide 2 is handled by code (title image). DO NOT create it in JSON.
    3) Slide 3 must be kind="key_points" titled "Key Points" and MUST use bullets summarizing the ENTIRE report.
    4) After that: for EACH H2 section in order:
       - Create ONE slide kind="section_text" with the section title and the section text as PARAGRAPHS (continuous text).
       - Do NOT summarize. Keep the wording as close to the original section_text as possible.
       - If the section_text is too long to fit, split into multiple section_text slides (1/2), (2/2) preserving original order.
       - Then, if that section has an image, create ONE slide kind="section_image" (image only) using the FIRST image from section_images.
    5) NEVER put text and image on the same slide (except title slide which has text only).
    6) Use layout capacity and choose layouts appropriately:
       - section_text should use TEXT layouts
       - section_image should use layouts that allow image-only (or blank)


    AVAILABLE TEMPLATE LAYOUTS (with text capacity limits):
    {layout_context}

    AVAILABLE IMAGES:
    {", ".join(image_list) if image_list else "No images available"}

    RETURN FORMAT (JSON only):
    {{
      "slides": [
        {{
          "kind": "title|image_only|key_points|section_text|section_image",
          "title": "slide title",
          "text": "paragraph text or null",
          "bullets": ["bullet 1", "bullet 2"],
          "image_path": "path/to/image.png or null",
          "layout_index": 0
        }}
      ]
    }}

    REPORT SECTIONS (H2):
    {sections_json}
    """

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are a PowerPoint expert. Output ONLY valid JSON. Respect layout capacity limits but use the available space effectively. Create comprehensive, informative slides."},
            {"role": "user", "content": prompt}
        ],
        response_format={"type": "json_object"}
    )
    return json.loads(response.choices[0].message.content).get('slides', [])

def fit_image_to_placeholder(placeholder, image_path):
    """Inserts image while maintaining aspect ratio within placeholder bounds."""
    from PIL import Image
    
    ph_width = placeholder.width
    ph_height = placeholder.height
    ph_left = placeholder.left
    ph_top = placeholder.top
    
    with Image.open(image_path) as img:
        img_width, img_height = img.size
        img_ratio = img_width / img_height
        ph_ratio = ph_width / ph_height
        
        if img_ratio > ph_ratio:
            new_width = ph_width
            new_height = int(ph_width / img_ratio)
        else:
            new_height = ph_height
            new_width = int(ph_height * img_ratio)
        
        left = ph_left + (ph_width - new_width) // 2
        top = ph_top + (ph_height - new_height) // 2
        
        return left, top, new_width, new_height

def split_slide_if_needed(slide_info, layouts):
    """
    Solo divide slides basadas en BULLETS (key_points).
    Si no es una slide con bullets, NO divide.
    Además, hace bullets robusto si viene None.
    """
    kind = slide_info.get("kind", "")
    
    # ✅ Solo aplicamos split a slides con bullets (Key Points)
    if kind != "key_points":
        # Asegura que bullets no sea None para evitar errores aguas abajo
        if slide_info.get("bullets") is None:
            slide_info["bullets"] = []
        return [slide_info]

    bullets = slide_info.get("bullets") or []  # robusto: None -> []
    slide_info["bullets"] = bullets

    layout_idx = slide_info.get("layout_index", 1)

    # Capacidad según layout (si existe)
    if layout_idx < len(layouts) and layouts[layout_idx].get("text_capacity"):
        capacity = layouts[layout_idx]["text_capacity"]
        max_bullets = capacity["max_bullets"]
        max_chars = capacity["max_chars_per_bullet"]
    else:
        max_bullets = 6
        max_chars = 80

    # Solo split si excede claramente capacidad
    needs_split = len(bullets) > int(max_bullets * 1.2)

    for b in bullets:
        if len(str(b)) > int(max_chars * 1.3):
            needs_split = True
            break

    if not needs_split:
        return [slide_info]

    # ---- Split real ----
    result_slides = []
    chunks = []
    current_chunk = []

    for bullet in bullets:
        bullet_text = str(bullet)

        # Si bullet es muy largo, intenta trocearlo por frases
        if len(bullet_text) > int(max_chars * 1.3) and ". " in bullet_text:
            sentences = bullet_text.split(". ")
            for sent in sentences:
                sent = sent.strip()
                if not sent:
                    continue
                if not sent.endswith("."):
                    sent += "."
                if len(current_chunk) >= max_bullets:
                    chunks.append(current_chunk)
                    current_chunk = [sent]
                else:
                    current_chunk.append(sent)
        else:
            if len(current_chunk) >= max_bullets:
                chunks.append(current_chunk)
                current_chunk = [bullet_text]
            else:
                current_chunk.append(bullet_text)

    if current_chunk:
        chunks.append(current_chunk)

    base_title = slide_info.get("title", "Slide")
    total_parts = len(chunks)

    for idx, chunk in enumerate(chunks):
        new_slide = slide_info.copy()
        if total_parts > 1:
            clean_title = re.sub(r"\s*\(\d+/\d+\)\s*$", "", base_title)
            new_slide["title"] = f"{clean_title} ({idx+1}/{total_parts})"
        new_slide["bullets"] = chunk
        # key_points no debería llevar imagen
        new_slide["image_path"] = None
        result_slides.append(new_slide)

    return result_slides

def add_fullscreen_image(slide, prs, image_path, margin_in=0.3):
    """
    Inserta una imagen centrada, manteniendo aspecto, ocupando casi toda la slide.
    """
    from PIL import Image

    slide_w = prs.slide_width
    slide_h = prs.slide_height

    margin = Inches(margin_in)
    max_w = slide_w - 2 * margin
    max_h = slide_h - 2 * margin

    with Image.open(image_path) as im:
        iw, ih = im.size
        img_ratio = iw / ih
        box_ratio = max_w / max_h

        if img_ratio > box_ratio:
            # limita por ancho
            w = max_w
            h = int(max_w / img_ratio)
        else:
            # limita por alto
            h = max_h
            w = int(max_h * img_ratio)

    left = int((slide_w - w) / 2)
    top = int((slide_h - h) / 2)
    slide.shapes.add_picture(image_path, left, top, width=w, height=h)

def add_paragraph_textbox(slide, prs, title, text, margin_in=0.6):
    """
    Text slide: título + texto en párrafos (sin bullets).
    """
    if slide.shapes.title:
        slide.shapes.title.text = title or ""
        if slide.shapes.title.has_text_frame:
            auto_fit_text(slide.shapes.title.text_frame, max_font_size=32, min_font_size=20)

    slide_w = prs.slide_width
    slide_h = prs.slide_height
    margin = Inches(margin_in)

    left = margin
    top = Inches(1.6)  # debajo del título
    width = slide_w - 2 * margin
    height = slide_h - top - margin

    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True

    paragraphs = [p.strip() for p in (text or "").split("\n\n") if p.strip()]
    if not paragraphs and (text or "").strip():
        paragraphs = [(text or "").strip()]

    for i, para in enumerate(paragraphs):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = para
        p.level = 0
        # asegura sin bullets
        try:
            p._p.get_or_add_pPr().get_or_add_buNone()
        except Exception:
            pass

    auto_fit_text(tf, max_font_size=18, min_font_size=10)

def create_presentation(slides_data, template_path, output_path, md_dir, layouts):
    """Constructs the presentation with improved text fitting and image placement."""
    prs = Presentation(template_path)

    num_original_slides = len(prs.slides)
    print(f"Template contains {num_original_slides} sample slides. They will be removed after generation.")

    # 1) Pre-procesar: split si hace falta
    processed_slides = []
    for slide_info in slides_data:
        split_slides = split_slide_if_needed(slide_info, layouts)
        processed_slides.extend(split_slides)

    if len(processed_slides) > len(slides_data):
        print(f"Note: Split {len(slides_data)} slides into {len(processed_slides)} slides for better readability")

    # 2) Crear slides
    for i, slide_info in enumerate(processed_slides):
        l_idx = slide_info.get("layout_index", 1)
        if l_idx >= len(prs.slide_layouts):
            l_idx = 1

        layout = prs.slide_layouts[l_idx]
        slide = prs.slides.add_slide(layout)

        # ------------------------------------------------------
        # A) Slide SOLO IMAGEN (pantalla completa)
        # ------------------------------------------------------
        if slide_info.get("image_only") and slide_info.get("image_path"):
            img_rel = slide_info["image_path"]
            full_img_path = os.path.normpath(os.path.join(md_dir, img_rel))

            # Quitar título si existe
            if slide.shapes.title:
                slide.shapes.title.text = ""

            # Eliminar placeholders para que no tapen
            for shp in list(slide.placeholders):
                try:
                    sp = shp._element
                    sp.getparent().remove(sp)
                except Exception:
                    pass

            if os.path.exists(full_img_path):
                add_fullscreen_image(slide, prs, full_img_path, margin_in=0.2)

            continue

        # ------------------------------------------------------
        # B) Slide según "kind"
        # ------------------------------------------------------
        kind = slide_info.get("kind", "section_text")
        title = slide_info.get("title", "Untitled")
        print(f"Creating Slide {i+1}: {title} [{kind}]")

        # ------------------------------------------------------
        # 1) KEY POINTS: bullets (solo aquí usamos bullets)
        # ------------------------------------------------------
        if kind == "key_points":
            if slide.shapes.title:
                slide.shapes.title.text = title
                if slide.shapes.title.has_text_frame:
                    auto_fit_text(slide.shapes.title.text_frame, max_font_size=32, min_font_size=20)

            text_placeholder = next(
                (s for s in slide.placeholders if s.placeholder_format.type in
                 [PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.SUBTITLE, PP_PLACEHOLDER.OBJECT]),
                None
            )

            if text_placeholder:
                tf = text_placeholder.text_frame
                tf.clear()
                tf.word_wrap = True

                for bullet_text in slide_info.get("bullets", []):
                    p = tf.add_paragraph()
                    p.text = str(bullet_text)
                    p.level = 0

                auto_fit_text(tf, max_font_size=18, min_font_size=10)

            continue

        # ------------------------------------------------------
        # 2) SECTION TEXT: texto en párrafos (SIN bullets)
        # ------------------------------------------------------
        if kind == "section_text":
            add_paragraph_textbox(slide, prs, title, slide_info.get("text", ""))
            continue

        # ------------------------------------------------------
        # 3) SECTION IMAGE: imagen fullscreen (sin texto)
        # ------------------------------------------------------
        if kind == "section_image":
            img_rel = slide_info.get("image_path")
            if img_rel:
                full_img_path = os.path.normpath(os.path.join(md_dir, img_rel))

                if slide.shapes.title:
                    slide.shapes.title.text = ""

                for shp in list(slide.placeholders):
                    try:
                        sp = shp._element
                        sp.getparent().remove(sp)
                    except Exception:
                        pass

                if os.path.exists(full_img_path):
                    add_fullscreen_image(slide, prs, full_img_path, margin_in=0.2)

            continue

        # ------------------------------------------------------
        # 4) Fallback: si llega algo raro, tratamos como texto
        # ------------------------------------------------------
        add_paragraph_textbox(slide, prs, title, slide_info.get("text", ""))

    # 3) Limpiar slides “de muestra” del template
    print(f"Cleaning up {num_original_slides} sample slides...")
    for _ in range(num_original_slides):
        delete_slide(prs, 0)

    # 4) Guardar
    try:
        prs.save(output_path)
        print(f"\nSUCCESS: Created presentation at {output_path}")
    except PermissionError:
        print(f"\nERROR: Cannot save to '{output_path}'")
        print("The file may have been opened during processing.")
        print("Please close it and run the script again.")
        sys.exit(1)

def validate_files(md_path, pptx_path):
    if not os.path.exists(md_path):
        print(f"Error: Markdown file not found: {md_path}")
        sys.exit(1)
    if not os.path.exists(pptx_path):
        print(f"Error: PowerPoint template not found: {pptx_path}")
        sys.exit(1)

def check_output_writable(output_path):
    """Check if output file can be written (not open in another program)."""
    if os.path.exists(output_path):
        try:
            with open(output_path, 'a'):
                pass
            return True
        except PermissionError:
            return False
    return True

def main():
    parser = argparse.ArgumentParser(description="Convert Markdown to PowerPoint with AI-powered layout selection")
    parser.add_argument("--md", required=True, help="Path to input Markdown file")
    parser.add_argument("--pptx", required=True, help="Path to PowerPoint template")
    parser.add_argument("--api_key", help="OpenAI API key (or set OPENAI_API_KEY env var)")
    parser.add_argument("--output", help="Output file path (default: adds '_output' to template name)")
    args = parser.parse_args()

    validate_files(args.md, args.pptx)
    
    api_key = args.api_key or os.getenv("OPENAI_API_KEY")
    if not api_key:
        print("Error: OpenAI API key required. Use --api_key or set OPENAI_API_KEY environment variable.")
        sys.exit(1)
    
    md_abs = os.path.abspath(args.md)
    md_dir = os.path.dirname(md_abs)
    
    if args.output:
        output_path = args.output
    else:
        base, ext = os.path.splitext(args.pptx)
        output_path = f"{base}_output{ext}"
    
    if not check_output_writable(output_path):
        print(f"\n❌ ERROR: Cannot write to '{output_path}'")
        print("The file may be open in PowerPoint or another program.")
        print("Please close the file and try again.\n")
        sys.exit(1)

    with open(md_abs, 'r', encoding='utf-8') as f:
        content = f.read()

    sections, images, title_h1 = parse_markdown(content)
    layouts = get_template_layouts(args.pptx)
    
    print("Analyzing content and generating slide structure...")
    slides_json = summarize_and_map_with_llm(sections, layouts, images, api_key, title_h1=title_h1)

    title_img = extract_title_image(content)

    if title_img:
        # 1) evita que el LLM la use en otra slide
        for s in slides_json:
            if s.get("image_path") == title_img:
                s["image_path"] = None

        # 2) asegura que la portada no tenga imagen
        if slides_json:
            slides_json[0]["image_path"] = None

        # 3) elige un layout "lo más vacío posible"
        #    prioridad: Blank -> layout sin placeholders -> cualquiera con [IMAGES]
        def pick_image_layout(layouts):
            for l in layouts:
                if "blank" in l["name"].lower():
                    return l["index"]
            for l in layouts:
                if l["placeholder_count"] == 0:
                    return l["index"]
            for l in layouts:
                ph_types = [p["type"] for p in l["placeholders"]]
                if any("PICTURE" in t for t in ph_types):
                    return l["index"]
            return 1  # fallback seguro

        img_layout = pick_image_layout(layouts)

        # 4) inserta slide 2 (índice 1)
        image_only_slide = {
            "title": "",              # sin título
            "bullets": [],            # sin bullets
            "image_path": title_img,  # la imagen del md
            "layout_index": img_layout,
            "notes": "Title image - full slide",
            "image_only": True
        }

        slides_json.insert(1, image_only_slide)

    
    print(f"Creating presentation with {len(slides_json)} slides...")
    create_presentation(slides_json, args.pptx, output_path, md_dir, layouts)

if __name__ == "__main__":
    main()
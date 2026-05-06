from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import StreamingResponse
from pptx import Presentation
import io

app = FastAPI()

# A simple endpoint to check if the server is awake
@app.get("/")
def health_check():
    return {"status": "PPTX Builder Service is online and ready!"}

@app.post("/generate-pptx")
async def generate_pptx(request: Request):
    try:
        # 1. Parse the JSON sent from n8n
        payload = await request.json()
        presentation_data = payload.get("presentation", {})
        slides_data = presentation_data.get("slides", [])
        
        # 2. Initialize a blank presentation
        prs = Presentation()
        
        # 3. Build the slides dynamically
        for slide_info in slides_data:
            layout_type = slide_info.get("layout", "title_slide")
            
            # Layout 0: Title Slide
            if layout_type == "title_slide":
                slide = prs.slides.add_slide(prs.slide_layouts[0])
                if slide.shapes.title:
                    slide.shapes.title.text = slide_info.get("title", "Presentation Title")
                if len(slide.placeholders) > 1:
                    slide.placeholders[1].text = slide_info.get("subtitle", "")
                    
            # Layout 1: Title and Bullets
            elif layout_type == "title_and_bullets" or layout_type == "split_right_image":
                slide = prs.slides.add_slide(prs.slide_layouts[1])
                if slide.shapes.title:
                    slide.shapes.title.text = slide_info.get("title", "")
                
                if len(slide.placeholders) > 1:
                    body = slide.placeholders[1]
                    tf = body.text_frame
                    bullets = slide_info.get("bullets", [])
                    
                    for i, bullet in enumerate(bullets):
                        if i == 0:
                            tf.text = bullet # The first bullet replaces default text
                        else:
                            p = tf.add_paragraph()
                            p.text = bullet
            else:
                # Fallback: Blank Slide
                slide = prs.slides.add_slide(prs.slide_layouts[6])

        # 4. Save to a virtual memory stream
        ppt_stream = io.BytesIO()
        prs.save(ppt_stream)
        ppt_stream.seek(0)
        
        # 5. Return the binary file
        return StreamingResponse(
            ppt_stream, 
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": "attachment; filename=Dynamic_Presentation.pptx"}
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
        
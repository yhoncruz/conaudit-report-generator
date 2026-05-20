import win32com.client
import os

def main():
    base_dir = os.path.abspath(os.path.dirname(__file__))
    pptx_path = os.path.join(base_dir, 'Presentación Reunión de Seguimiento.pptx')
    
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    try:
        pres = ppt.Presentations.Open(pptx_path, WithWindow=False)
        slide = pres.Slides(1)
        print(f"Total shapes in Slide 1: {slide.Shapes.Count}")
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            print(f"Shape {i}: Name='{shape.Name}', Type={shape.Type}, Width={shape.Width}, Height={shape.Height}, Top={shape.Top}, Left={shape.Left}")
            if shape.HasTextFrame:
                try:
                    text = shape.TextFrame.TextRange.Text
                    print(f"  Text preview: {text[:30]}")
                except:
                    pass
    except Exception as e:
        print(f"Error: {e}")
    finally:
        try:
            pres.Close()
        except:
            pass
        ppt.Quit()

if __name__ == "__main__":
    main()

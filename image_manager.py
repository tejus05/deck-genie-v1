import io
from typing import Dict, Optional
from PIL import Image

class ImageManager:
    """Manages custom images for slides."""
    
    def __init__(self, uploaded_images: Dict[str, bytes]):
        self.uploaded_images = uploaded_images
    
    def get_image_for_slide(self, slide_key: str) -> Optional[io.BytesIO]:
        """Get custom image for a slide if available."""
        if slide_key in self.uploaded_images:
            image_data = self.uploaded_images[slide_key]
            # Process and return the image
            return self._process_image(image_data)
        return None
    
    def _process_image(self, image_data: bytes) -> io.BytesIO:
        """Process uploaded image to ensure compatibility."""
        try:
            # Open the image
            image = Image.open(io.BytesIO(image_data))
            
            # Convert to RGB if necessary (removes alpha channel)
            if image.mode in ('RGBA', 'LA', 'P'):
                image = image.convert('RGB')
            
            # Resize if too large (max 1920x1080)
            max_width, max_height = 1920, 1080
            if image.width > max_width or image.height > max_height:
                image.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
            
            # Save processed image to BytesIO
            output = io.BytesIO()
            image.save(output, format='JPEG', quality=85)
            output.seek(0)
            
            return output
            
        except Exception as e:
            print(f"Error processing image: {str(e)}")
            return None
    
    def has_custom_image(self, slide_key: str) -> bool:
        """Check if a slide has a custom image."""
        return slide_key in self.uploaded_images

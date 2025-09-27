"""Shape helpers for premium Slides generation."""
from __future__ import annotations

from typing import Any, Dict, List, Optional


class SlidesShapeHelper:
    def __init__(self, slides_client):
        self.client = slides_client
        
    def add_rectangle(
        self,
        presentation_id: str,
        page_id: str,
        object_id: str,
        x: int,
        y: int,
        w: int,
        h: int,
        fill_hex: Optional[str] = None,
        border_hex: Optional[str] = None,
    ) -> None:
        """Add filled rectangle shape."""
        requests = [{
            "createShape": {
                "objectId": object_id,
                "shapeType": "RECTANGLE",
                "elementProperties": {
                    "pageObjectId": page_id,
                    "size": {"width": {"magnitude": w, "unit": "PT"}, "height": {"magnitude": h, "unit": "PT"}},
                    "transform": {"scaleX": 1, "scaleY": 1, "translateX": x, "translateY": y, "unit": "PT"}
                }
            }
        }]
        
        if fill_hex:
            requests.append({
                "updateShapeProperties": {
                    "objectId": object_id,
                    "fields": "shapeBackgroundFill.solidFill.color",
                    "shapeProperties": {
                        "shapeBackgroundFill": {"solidFill": {"color": {"rgbColor": self._hex_to_rgb01(fill_hex)}}}
                    }
                }
            })
            
        if border_hex:
            requests.append({
                "updateShapeProperties": {
                    "objectId": object_id,
                    "fields": "outline",
                    "shapeProperties": {
                        "outline": {
                            "outlineFill": {"solidFill": {"color": {"rgbColor": self._hex_to_rgb01(border_hex)}}},
                            "weight": {"magnitude": 1, "unit": "PT"}
                        }
                    }
                }
            })
        
        self.client.presentations().batchUpdate(
            presentationId=presentation_id, body={"requests": requests}
        ).execute()
        
    def _hex_to_rgb01(self, hex_color: str) -> Dict[str, float]:
        try:
            hex_color = hex_color.lstrip('#')
            r = int(hex_color[0:2], 16) / 255.0
            g = int(hex_color[2:4], 16) / 255.0
            b = int(hex_color[4:6], 16) / 255.0
            return {"red": r, "green": g, "blue": b}
        except Exception:
            return {"red": 0.8, "green": 0.0, "blue": 0.0}

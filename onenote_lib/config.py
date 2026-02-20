from pydantic_settings import BaseSettings


class OneNoteConfig(BaseSettings):
    model_config = {"env_prefix": "ONENOTE_"}

    vision_url: str = "http://localhost:1234"
    vision_model: str = ""
    vision_fallback_url: str = ""
    vision_fallback_model: str = ""
    max_image_size_kb: int = 512
    max_images_per_page: int = 20


config = OneNoteConfig()

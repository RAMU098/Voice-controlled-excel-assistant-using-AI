"""
Speech Handler Module
Handles text-to-speech conversion for voice feedback
"""
import os
import tempfile
import base64
from typing import Optional


class SpeechHandler:
    """
    Handles text-to-speech conversion for providing audio feedback.
    Uses gTTS (Google Text-to-Speech) for generating audio.
    """
    
    def __init__(self):
        """Initialize the speech handler."""
        self.gtts_available = False
        self.pyttsx3_available = False
        
        # Try to initialize gTTS
        try:
            from gtts import gTTS
            self.gtts_available = True
        except ImportError:
            pass
        
        # Try to initialize pyttsx3 as fallback
        try:
            import pyttsx3
            self.pyttsx3_available = True
        except ImportError:
            pass
    
    def text_to_speech_base64(self, text: str, lang: str = 'en') -> Optional[str]:
        """
        Convert text to speech and return as base64 encoded audio.
        
        Args:
            text: Text to convert to speech
            lang: Language code (default: 'en')
            
        Returns:
            Base64 encoded MP3 audio or None if conversion fails
        """
        if not self.gtts_available:
            return None
        
        try:
            from gtts import gTTS
            import io
            
            # Create gTTS object
            tts = gTTS(text=text, lang=lang)
            
            # Save to bytes buffer
            audio_buffer = io.BytesIO()
            tts.write_to_fp(audio_buffer)
            audio_buffer.seek(0)
            
            # Encode to base64
            audio_base64 = base64.b64encode(audio_buffer.read()).decode('utf-8')
            
            return audio_base64
        except Exception as e:
            print(f"TTS error: {e}")
            return None
    
    def text_to_speech_file(self, text: str, output_path: str, lang: str = 'en') -> bool:
        """
        Convert text to speech and save to a file.
        
        Args:
            text: Text to convert to speech
            output_path: Path to save the audio file
            lang: Language code (default: 'en')
            
        Returns:
            True if successful, False otherwise
        """
        if not self.gtts_available:
            return False
        
        try:
            from gtts import gTTS
            
            tts = gTTS(text=text, lang=lang)
            tts.save(output_path)
            return True
        except Exception as e:
            print(f"TTS file save error: {e}")
            return False
    
    def get_audio_data_uri(self, text: str, lang: str = 'en') -> Optional[str]:
        """
        Get a data URI for the audio that can be used directly in HTML.
        
        Args:
            text: Text to convert to speech
            lang: Language code
            
        Returns:
            Data URI string or None
        """
        audio_base64 = self.text_to_speech_base64(text, lang)
        if audio_base64:
            return f"data:audio/mp3;base64,{audio_base64}"
        return None


# Singleton instance
_speech_handler = None

def get_speech_handler() -> SpeechHandler:
    """Get the singleton SpeechHandler instance."""
    global _speech_handler
    if _speech_handler is None:
        _speech_handler = SpeechHandler()
    return _speech_handler


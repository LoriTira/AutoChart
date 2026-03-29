"""Abstract base class for worksheet parsers."""

from abc import ABC, abstractmethod
from autochart.config import ChartConfig


class BaseParser(ABC):
    """Base class for all worksheet parsers."""

    @abstractmethod
    def parse(self, worksheet, config: ChartConfig) -> dict:
        """Parse a worksheet and return structured data.

        Args:
            worksheet: An openpyxl worksheet object.
            config: Chart configuration with disease name, demographics, etc.

        Returns:
            A dict mapping ChartSetType to the parsed data object(s).
        """
        pass

    @abstractmethod
    def can_parse(self, worksheet) -> bool:
        """Check if this parser can handle the given worksheet.

        Args:
            worksheet: An openpyxl worksheet object.

        Returns:
            True if this parser can handle the worksheet format.
        """
        pass

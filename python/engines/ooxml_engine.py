"""OOXMLEngine: assembled from the per-domain mixins in engines/ooxml/."""

from __future__ import annotations

from engines.ooxml.charts import OOXMLChartsMixin
from engines.ooxml.core import OOXMLCore
from engines.ooxml.discovery import OOXMLDiscoveryMixin
from engines.ooxml.session import OOXMLSessionMixin
from engines.ooxml.shapes import OOXMLShapesMixin
from engines.ooxml.slides import OOXMLSlidesMixin
from engines.ooxml.snapshot import OOXMLSnapshotMixin
from engines.ooxml.tables import OOXMLTablesMixin
from engines.ooxml.text import OOXMLTextMixin


class OOXMLEngine(
    OOXMLSessionMixin,
    OOXMLDiscoveryMixin,
    OOXMLSlidesMixin,
    OOXMLSnapshotMixin,
    OOXMLTextMixin,
    OOXMLShapesMixin,
    OOXMLTablesMixin,
    OOXMLChartsMixin,
    OOXMLCore,
):
    """Public engine. Method resolution: domain mixins first, then the shared OOXMLCore."""

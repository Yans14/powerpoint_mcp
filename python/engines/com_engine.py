"""COMEngine: assembled from the per-domain mixins in engines/com/."""

from __future__ import annotations

from engines.com.charts import COMChartsMixin
from engines.com.core import COMCore
from engines.com.discovery import COMDiscoveryMixin
from engines.com.session import COMSessionMixin
from engines.com.shapes import COMShapesMixin
from engines.com.slides import COMSlidesMixin
from engines.com.snapshot import COMSnapshotMixin
from engines.com.tables import COMTablesMixin
from engines.com.text import COMTextMixin


class COMEngine(
    COMSessionMixin,
    COMDiscoveryMixin,
    COMSlidesMixin,
    COMSnapshotMixin,
    COMTextMixin,
    COMShapesMixin,
    COMTablesMixin,
    COMChartsMixin,
    COMCore,
):
    """Public engine. Method resolution: domain mixins first, then the shared COMCore."""

"""|DocumentPart| and closely related objects."""

from __future__ import annotations

import os
from typing import IO, TYPE_CHECKING, cast

from docx.document import Document
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.parts.comments import CommentsPart
from docx.parts.hdrftr import FooterPart, HeaderPart
from docx.parts.numbering import NumberingPart
from docx.parts.settings import SettingsPart
from docx.parts.story import StoryPart
from docx.parts.styles import StylesPart
from docx.shape import InlineShapes
from docx.shared import lazyproperty

if TYPE_CHECKING:
    from docx.comments import Comments
    from docx.enum.style import WD_STYLE_TYPE
    from docx.opc.coreprops import CoreProperties
    from docx.settings import Settings
    from docx.styles.style import BaseStyle


class DocumentPart(StoryPart):
    """Main document part of a WordprocessingML (WML) package, aka a .docx file.

    Acts as broker to other parts such as image, core properties, and style parts. It
    also acts as a convenient delegate when a mid-document object needs a service
    involving a remote ancestor. The `Parented.part` property inherited by many content
    objects provides access to this part object for that purpose.
    """

    def add_footer_part(self):
        """Return (footer_part, rId) pair for newly-created footer part."""
        footer_part = FooterPart.new(self.package)
        rId = self.relate_to(footer_part, RT.FOOTER)
        return footer_part, rId

    def add_header_part(self):
        """Return (header_part, rId) pair for newly-created header part."""
        header_part = HeaderPart.new(self.package)
        rId = self.relate_to(header_part, RT.HEADER)
        return header_part, rId

    @property
    def comments(self) -> Comments:
        """|Comments| object providing access to the comments added to this document."""
        return self._comments_part.comments

    @property
    def core_properties(self) -> CoreProperties:
        """A |CoreProperties| object providing read/write access to the core properties
        of this document."""
        return self.package.core_properties

    @property
    def document(self):
        """A |Document| object providing access to the content of this document."""
        return Document(self._element, self)

    def drop_header_part(self, rId: str) -> None:
        """Remove related header part identified by `rId`."""
        self.drop_rel(rId)

    def footer_part(self, rId: str):
        """Return |FooterPart| related by `rId`."""
        return self.related_parts[rId]

    def get_style(self, style_id: str | None, style_type: WD_STYLE_TYPE) -> BaseStyle:
        """Return the style in this document matching `style_id`.

        Returns the default style for `style_type` if `style_id` is |None| or does not
        match a defined style of `style_type`.
        """
        return self.styles.get_by_id(style_id, style_type)

    def get_style_id(self, style_or_name, style_type):
        """Return the style_id (|str|) of the style of `style_type` matching
        `style_or_name`.

        Returns |None| if the style resolves to the default style for `style_type` or if
        `style_or_name` is itself |None|. Raises if `style_or_name` is a style of the
        wrong type or names a style not present in the document.
        """
        return self.styles.get_style_id(style_or_name, style_type)

    def header_part(self, rId: str):
        """Return |HeaderPart| related by `rId`."""
        return self.related_parts[rId]

    @lazyproperty
    def inline_shapes(self):
        """The |InlineShapes| instance containing the inline shapes in the document."""
        return InlineShapes(self._element.body, self)

    @lazyproperty
    def numbering_part(self) -> NumberingPart:
        """A |NumberingPart| object providing access to the numbering definitions for this document.

        Creates an empty numbering part if one is not present.
        """
        try:
            return cast(NumberingPart, self.part_related_by(RT.NUMBERING))
        except KeyError:
            numbering_part = NumberingPart.new()
            self.relate_to(numbering_part, RT.NUMBERING)
            return numbering_part

    def save(self, path_or_stream: str | IO[bytes], preserve_macros: bool | None = None):
        """Save this document to `path_or_stream`, which can be either a path to a
        filesystem location (a string) or a file-like object.

        Args:
            path_or_stream: File path (string) or file-like object to save to
            preserve_macros: If True, preserves VBA macros when saving DOCM files.
                           If False, strips macros and converts to DOCX format.
                           If None (default), auto-detects based on file extension:
                           - .docm extension -> preserve macros
                           - .docx extension or no extension -> strip macros

        Note:
            When stripping macros from a DOCM file, if saving to a path with .docm
            extension, the extension will be automatically changed to .docx to match
            the content type and ensure Word can open the file correctly.
        """
        # Determine whether to preserve macros
        should_preserve = self._should_preserve_macros(path_or_stream, preserve_macros)

        # Only strip macros if we have a macro-enabled document and shouldn't preserve
        will_strip_macros = (
            self._content_type == CT.WML_DOCUMENT_MACRO_ENABLED_MAIN and not should_preserve
        )

        if will_strip_macros:
            # Convert to standard Word document type
            self._content_type = CT.WML_DOCUMENT_MAIN
            # Remove VBA and ActiveX control relationships and parts
            self._remove_macro_relationships()
            self._remove_vba_parts()

            # If saving to a file path with .docm extension, correct it to .docx
            # to match the content type (Word won't open files with mismatched extensions)
            if isinstance(path_or_stream, str) and path_or_stream.lower().endswith('.docm'):
                path_or_stream = path_or_stream[:-5] + '.docx'

        self.package.save(path_or_stream)

    def _should_preserve_macros(
        self, path_or_stream: str | IO[bytes], preserve_macros: bool | None
    ) -> bool:
        """Determine whether to preserve VBA macros based on file extension or explicit flag.

        Args:
            path_or_stream: Target file path or stream
            preserve_macros: Explicit preserve flag (overrides auto-detection)

        Returns:
            True if macros should be preserved, False if they should be stripped
        """
        # If explicitly specified, use that value
        if preserve_macros is not None:
            return preserve_macros

        # For file-like objects, default to stripping (safer)
        if not isinstance(path_or_stream, str):
            return False

        # For file paths, detect based on extension
        _, ext = os.path.splitext(path_or_stream)
        return ext.lower() == ".docm"

    def _remove_macro_relationships(self):
        """Remove VBA project and ActiveX control relationships from the document part."""
        # Relationship types that should be removed when converting from DOCM to DOCX
        macro_reltypes = (
            "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
            "http://schemas.microsoft.com/office/2006/relationships/wordVbaData",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control",
        )
        # Find and remove relationships of these types
        rids_to_remove = [
            rel.rId for rel in self.rels.values()
            if rel.reltype in macro_reltypes
        ]

        # Remove control elements from document XML that reference these relationships
        for rId in rids_to_remove:
            # Find and remove any <w:control> elements that reference this rId
            control_elements = self._element.xpath(f'.//w:control[@r:id="{rId}"]')
            for control_elem in control_elements:
                control_elem.getparent().remove(control_elem)

            # Remove the relationship
            del self.rels[rId]

    def _remove_vba_parts(self):
        """Remove VBA-related parts from the package.

        This removes the actual binary VBA parts and other macro-related parts,
        not just the relationships to them.
        """
        assert self.package is not None

        # VBA-related content types to remove
        vba_content_types = (
            "application/vnd.ms-word.vbaProject",
            "application/vnd.ms-word.vbaData+xml",
            "application/vnd.ms-office.activeX",
            "application/vnd.ms-office.activeX+xml",
        )

        # VBA-related partname patterns
        vba_partname_patterns = (
            "/word/vbaProject.bin",
            "/word/vbaData.xml",
        )

        # Collect parts to remove
        parts_to_remove = []
        for part in self.package.parts:
            # Check by content type
            if part.content_type in vba_content_types:
                parts_to_remove.append(part)
            # Check by partname pattern
            elif any(str(part.partname) == pattern for pattern in vba_partname_patterns):
                parts_to_remove.append(part)

        # Remove the parts from package
        # Note: We need to remove them from the package's internal tracking
        # This is tricky because the package builds its parts list dynamically via iter_parts()
        # We'll remove relationships pointing to these parts from all sources
        for part_to_remove in parts_to_remove:
            # Remove relationships from package level
            for rel in list(self.package.rels.values()):
                if not rel.is_external and rel.target_part == part_to_remove:
                    del self.package.rels[rel.rId]

            # Remove relationships from all other parts
            for part in self.package.parts:
                if part == part_to_remove:
                    continue
                for rel in list(part.rels.values()):
                    if not rel.is_external and rel.target_part == part_to_remove:
                        del part.rels[rel.rId]

    @property
    def settings(self) -> Settings:
        """A |Settings| object providing access to the settings in the settings part of
        this document."""
        return self._settings_part.settings

    @property
    def styles(self):
        """A |Styles| object providing access to the styles in the styles part of this
        document."""
        return self._styles_part.styles

    @property
    def _comments_part(self) -> CommentsPart:
        """A |CommentsPart| object providing access to the comments added to this document.

        Creates a default comments part if one is not present.
        """
        try:
            return cast(CommentsPart, self.part_related_by(RT.COMMENTS))
        except KeyError:
            assert self.package is not None
            comments_part = CommentsPart.default(self.package)
            self.relate_to(comments_part, RT.COMMENTS)
            return comments_part

    @property
    def _settings_part(self) -> SettingsPart:
        """A |SettingsPart| object providing access to the document-level settings for
        this document.

        Creates a default settings part if one is not present.
        """
        try:
            return cast(SettingsPart, self.part_related_by(RT.SETTINGS))
        except KeyError:
            settings_part = SettingsPart.default(self.package)
            self.relate_to(settings_part, RT.SETTINGS)
            return settings_part

    @property
    def _styles_part(self) -> StylesPart:
        """Instance of |StylesPart| for this document.

        Creates an empty styles part if one is not present.
        """
        try:
            return cast(StylesPart, self.part_related_by(RT.STYLES))
        except KeyError:
            package = self.package
            assert package is not None
            styles_part = StylesPart.default(package)
            self.relate_to(styles_part, RT.STYLES)
            return styles_part

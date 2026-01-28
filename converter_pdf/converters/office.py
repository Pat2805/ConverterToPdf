"""
Convertisseurs Microsoft Office (Word, Excel, PowerPoint).

Utilise COM avec DispatchEx pour créer des instances dédiées,
évitant les conflits avec les applications Office déjà ouvertes.
"""

from __future__ import annotations

import time
from pathlib import Path
from typing import TYPE_CHECKING

from .base import BaseConverter, ConversionResult, ConversionStatus
from ..com_utils import (
    WIN32COM_AVAILABLE,
    office_app_context,
    is_password_error,
    COMError,
    COMTimeoutError,
)

if TYPE_CHECKING:
    from ..config import Config
    from ..logger import ConverterLogger


class OfficeWordConverter(BaseConverter):
    """
    Convertisseur Word via Microsoft Office COM.

    Utilise Word.Application avec DispatchEx pour garantir
    une nouvelle instance (pas de conflit avec Word ouvert).
    """

    name = "office_word"
    supported_extensions = [".doc", ".docx", ".rtf", ".odt"]

    def is_available(self) -> bool:
        """Vérifie que pywin32 est installé."""
        return WIN32COM_AVAILABLE

    def convert(self, source: Path, dest: Path) -> ConversionResult:
        """Convertit un document Word en PDF via COM."""
        start = time.time()

        if not self.is_available():
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                message="pywin32 non installé",
            )

        try:
            with office_app_context(
                "Word.Application",
                logger=self.logger,
            ) as word:
                self.logger.debug("Word.Application créé (nouvelle instance)")

                # Ouvrir le document en lecture seule
                doc = word.Documents.Open(
                    str(source.absolute()),
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    ConfirmConversions=False,
                    NoEncodingDialog=True,
                    # Mot de passe invalide pour détecter les fichiers protégés
                    PasswordDocument="__INVALID__",
                    WritePasswordDocument="__INVALID__",
                    Revert=False,
                    Visible=False,
                )
                self.logger.debug(f"Document ouvert: {source.name}")

                try:
                    # Export PDF via ExportAsFixedFormat (méthode préférée)
                    # 17 = wdExportFormatPDF
                    doc.ExportAsFixedFormat(
                        OutputFileName=str(dest.absolute()),
                        ExportFormat=17,
                        OpenAfterExport=False,
                        OptimizeFor=0,      # wdExportOptimizeForPrint
                        Range=0,            # wdExportAllDocument
                        Item=0,             # wdExportDocumentContent
                        IncludeDocProps=True,
                        KeepIRM=True,
                        CreateBookmarks=1,  # wdExportCreateHeadingBookmarks
                        DocStructureTags=True,
                        BitmapMissingFonts=True,
                        UseISO19005_1=False,
                    )
                    self.logger.debug("ExportAsFixedFormat réussi")

                except Exception as export_err:
                    # Fallback: SaveAs2 avec FileFormat=17
                    self.logger.debug(f"ExportAsFixedFormat échoué: {export_err}, essai SaveAs2")
                    try:
                        if hasattr(doc, "SaveAs2"):
                            doc.SaveAs2(str(dest.absolute()), FileFormat=17)
                        else:
                            doc.SaveAs(str(dest.absolute()), FileFormat=17)
                        self.logger.debug("SaveAs2/SaveAs réussi")
                    except Exception as save_err:
                        raise export_err  # Relever l'erreur originale

                finally:
                    doc.Close(False)
                    self.logger.debug("Document fermé")

                duration = time.time() - start
                return ConversionResult(
                    status=ConversionStatus.SUCCESS,
                    source=source,
                    dest=dest,
                    duration=duration,
                    method=self.name,
                )

        except Exception as e:
            duration = time.time() - start

            if is_password_error(e):
                self.logger.warning(f"Document protégé par mot de passe: {source.name}")
                # Supprimer un éventuel PDF partiel
                if dest.exists():
                    try:
                        dest.unlink()
                    except Exception:
                        pass
                return ConversionResult(
                    status=ConversionStatus.SKIPPED_PASSWORD,
                    source=source,
                    dest=None,
                    duration=duration,
                    method=self.name,
                    message="Document protégé par mot de passe",
                    exception=e,
                )

            self.logger.error(f"Échec conversion Word: {e}", exc=e)
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=duration,
                method=self.name,
                exception=e,
            )


class OfficeExcelConverter(BaseConverter):
    """
    Convertisseur Excel via Microsoft Office COM.

    IMPORTANT: Utilise DispatchEx (pas Dispatch) pour éviter
    les conflits avec Excel déjà ouvert.
    """

    name = "office_excel"
    supported_extensions = [".xls", ".xlsx", ".xlsm", ".xlsb"]

    def is_available(self) -> bool:
        return WIN32COM_AVAILABLE

    def convert(self, source: Path, dest: Path) -> ConversionResult:
        """Convertit un classeur Excel en PDF via COM."""
        start = time.time()

        if not self.is_available():
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                message="pywin32 non installé",
            )

        try:
            with office_app_context(
                "Excel.Application",
                logger=self.logger,
            ) as excel:
                self.logger.debug("Excel.Application créé (nouvelle instance)")

                # Désactiver les mises à jour de liens
                try:
                    excel.AskToUpdateLinks = False
                except AttributeError:
                    pass

                # Ouvrir le classeur en lecture seule
                wb = excel.Workbooks.Open(
                    str(source.absolute()),
                    ReadOnly=True,
                    UpdateLinks=0,
                    Password="",
                    WriteResPassword="",
                    IgnoreReadOnlyRecommended=True,
                    AddToMru=False,
                )
                self.logger.debug(f"Classeur ouvert: {source.name}")

                try:
                    # Export PDF (Type=0 = xlTypePDF)
                    wb.ExportAsFixedFormat(
                        Type=0,
                        Filename=str(dest.absolute()),
                        Quality=0,  # xlQualityStandard
                        IncludeDocProperties=True,
                        IgnorePrintAreas=False,
                        OpenAfterPublish=False,
                    )
                    self.logger.debug("ExportAsFixedFormat réussi")
                finally:
                    wb.Close(False)
                    self.logger.debug("Classeur fermé")

                duration = time.time() - start
                return ConversionResult(
                    status=ConversionStatus.SUCCESS,
                    source=source,
                    dest=dest,
                    duration=duration,
                    method=self.name,
                )

        except Exception as e:
            duration = time.time() - start

            if is_password_error(e):
                self.logger.warning(f"Classeur protégé par mot de passe: {source.name}")
                if dest.exists():
                    try:
                        dest.unlink()
                    except Exception:
                        pass
                return ConversionResult(
                    status=ConversionStatus.SKIPPED_PASSWORD,
                    source=source,
                    dest=None,
                    duration=duration,
                    method=self.name,
                    message="Classeur protégé par mot de passe",
                    exception=e,
                )

            self.logger.error(f"Échec conversion Excel: {e}", exc=e)
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=duration,
                method=self.name,
                exception=e,
            )


class OfficePowerPointConverter(BaseConverter):
    """
    Convertisseur PowerPoint via Microsoft Office COM.

    Utilise DispatchEx pour créer une nouvelle instance.
    """

    name = "office_powerpoint"
    supported_extensions = [".ppt", ".pptx"]

    def is_available(self) -> bool:
        return WIN32COM_AVAILABLE

    def convert(self, source: Path, dest: Path) -> ConversionResult:
        """Convertit une présentation PowerPoint en PDF via COM."""
        start = time.time()

        if not self.is_available():
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=time.time() - start,
                method=self.name,
                message="pywin32 non installé",
            )

        try:
            with office_app_context(
                "PowerPoint.Application",
                logger=self.logger,
            ) as ppt:
                self.logger.debug("PowerPoint.Application créé (nouvelle instance)")

                # Ouvrir la présentation
                presentation = ppt.Presentations.Open(
                    str(source.absolute()),
                    ReadOnly=True,
                    Untitled=False,
                    WithWindow=False,
                )
                self.logger.debug(f"Présentation ouverte: {source.name}")

                try:
                    # Export PDF (32 = ppSaveAsPDF)
                    presentation.SaveAs(
                        str(dest.absolute()),
                        32,  # ppSaveAsPDF
                    )
                    self.logger.debug("SaveAs PDF réussi")
                finally:
                    presentation.Close()
                    self.logger.debug("Présentation fermée")

                duration = time.time() - start
                return ConversionResult(
                    status=ConversionStatus.SUCCESS,
                    source=source,
                    dest=dest,
                    duration=duration,
                    method=self.name,
                )

        except Exception as e:
            duration = time.time() - start

            if is_password_error(e):
                self.logger.warning(f"Présentation protégée: {source.name}")
                if dest.exists():
                    try:
                        dest.unlink()
                    except Exception:
                        pass
                return ConversionResult(
                    status=ConversionStatus.SKIPPED_PASSWORD,
                    source=source,
                    dest=None,
                    duration=duration,
                    method=self.name,
                    message="Présentation protégée par mot de passe",
                    exception=e,
                )

            self.logger.error(f"Échec conversion PowerPoint: {e}", exc=e)
            return ConversionResult(
                status=ConversionStatus.FAILED,
                source=source,
                dest=None,
                duration=duration,
                method=self.name,
                exception=e,
            )

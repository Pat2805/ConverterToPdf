"""
Rapport de session pour les conversions.

Génère un rapport texte lisible récapitulant une session de conversion,
avec statistiques détaillées, erreurs et fichiers problématiques.
"""

from __future__ import annotations

import traceback
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from .converters.base import ConversionResult, ConversionStatus


@dataclass
class FileStats:
    """Statistiques pour un type de fichier."""
    count: int = 0
    success: int = 0
    failed: int = 0
    skipped_exists: int = 0
    skipped_password: int = 0
    skipped_unsupported: int = 0
    source_size_bytes: int = 0
    dest_size_bytes: int = 0
    total_duration: float = 0.0


@dataclass
class SessionReport:
    """
    Rapport de session de conversion.

    Collecte les résultats de conversion et génère un rapport
    texte lisible à la fin de la session.
    """

    # Métadonnées de session
    start_time: datetime = field(default_factory=datetime.now)
    end_time: datetime | None = None
    source_directory: Path | None = None
    output_directory: Path | None = None
    recursive: bool = False

    # Résultats collectés
    results: list["ConversionResult"] = field(default_factory=list)

    # Statistiques par type de fichier
    stats_by_type: dict[str, FileStats] = field(default_factory=lambda: defaultdict(FileStats))

    # Listes pour le rapport détaillé
    conversions: list[tuple[Path, Path, str, float]] = field(default_factory=list)  # (source, dest, method, duration)
    errors: list[tuple[Path, str, str]] = field(default_factory=list)  # (path, reason, details)
    password_protected: list[Path] = field(default_factory=list)
    skipped_existing: list[Path] = field(default_factory=list)

    # Totaux
    total_source_bytes: int = 0
    total_dest_bytes: int = 0
    total_duration: float = 0.0

    def add_result(self, result: "ConversionResult") -> None:
        """
        Ajoute un résultat de conversion au rapport.

        Args:
            result: Résultat de la conversion
        """
        self.results.append(result)

        # Extension du fichier
        ext = result.source.suffix.lower()
        stats = self.stats_by_type[ext]
        stats.count += 1

        # Tailles
        source_size = int(result.source_size_mb * 1024 * 1024)
        dest_size = int(result.dest_size_mb * 1024 * 1024)
        stats.source_size_bytes += source_size
        self.total_source_bytes += source_size

        if result.dest and result.dest.exists():
            stats.dest_size_bytes += dest_size
            self.total_dest_bytes += dest_size

        # Durée
        stats.total_duration += result.duration
        self.total_duration += result.duration

        # Status
        status = result.status.value
        if status == "success":
            stats.success += 1
            # Collecter les conversions réussies
            dest_path = result.dest if result.dest else result.source.with_suffix(result.source.suffix + ".pdf")
            self.conversions.append((result.source, dest_path, result.method, result.duration))
        elif status == "failed":
            stats.failed += 1
            # Collecter les détails de l'erreur avec traceback complet
            error_msg = result.message or "Erreur inconnue"
            error_details = ""
            if result.exception:
                # Inclure le type d'exception et la traceback complète
                exc_type = type(result.exception).__name__
                exc_msg = str(result.exception)
                # Obtenir la traceback si disponible
                tb_str = "".join(traceback.format_exception(
                    type(result.exception),
                    result.exception,
                    result.exception.__traceback__
                ))
                error_details = f"{exc_type}: {exc_msg}\n{tb_str}"
            self.errors.append((result.source, error_msg, error_details))
        elif status == "skipped_exists":
            stats.skipped_exists += 1
            self.skipped_existing.append(result.source)
        elif status == "skipped_password":
            stats.skipped_password += 1
            self.password_protected.append(result.source)
        elif status == "skipped_unsupported":
            stats.skipped_unsupported += 1

    def finalize(self) -> None:
        """Finalise le rapport (appelé à la fin de la session)."""
        self.end_time = datetime.now()

    def _format_size(self, size_bytes: int) -> str:
        """Formate une taille en bytes en format lisible."""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        elif size_bytes < 1024 * 1024 * 1024:
            return f"{size_bytes / (1024 * 1024):.1f} MB"
        else:
            return f"{size_bytes / (1024 * 1024 * 1024):.2f} GB"

    def _format_duration(self, seconds: float) -> str:
        """Formate une durée en format lisible."""
        if seconds < 60:
            return f"{seconds:.1f}s"
        elif seconds < 3600:
            mins = int(seconds // 60)
            secs = seconds % 60
            return f"{mins}m {secs:.0f}s"
        else:
            hours = int(seconds // 3600)
            mins = int((seconds % 3600) // 60)
            return f"{hours}h {mins}m"

    @property
    def total_files(self) -> int:
        """Nombre total de fichiers traités."""
        return len(self.results)

    @property
    def total_success(self) -> int:
        """Nombre de conversions réussies."""
        return sum(s.success for s in self.stats_by_type.values())

    @property
    def total_failed(self) -> int:
        """Nombre d'échecs."""
        return sum(s.failed for s in self.stats_by_type.values())

    @property
    def total_skipped(self) -> int:
        """Nombre de fichiers ignorés."""
        return sum(
            s.skipped_exists + s.skipped_password + s.skipped_unsupported
            for s in self.stats_by_type.values()
        )

    def generate(self) -> str:
        """
        Génère le rapport texte complet.

        Returns:
            Rapport formaté en texte
        """
        lines = []
        sep = "=" * 80
        sep_light = "-" * 80

        # En-tête
        lines.append(sep)
        lines.append(f"RAPPORT DE CONVERSION - {self.start_time.strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(sep)
        lines.append("")

        # Informations de session
        lines.append("SESSION")
        lines.append(sep_light)
        if self.source_directory:
            lines.append(f"  Répertoire source  : {self.source_directory}")
        if self.output_directory:
            lines.append(f"  Répertoire sortie  : {self.output_directory}")
        lines.append(f"  Mode récursif      : {'Oui' if self.recursive else 'Non'}")
        if self.end_time:
            duration = (self.end_time - self.start_time).total_seconds()
            lines.append(f"  Durée session      : {self._format_duration(duration)}")
        lines.append("")

        # Résumé global
        lines.append("RÉSUMÉ")
        lines.append(sep_light)
        lines.append(f"  Fichiers analysés  : {self.total_files}")
        if self.total_files > 0:
            success_pct = (self.total_success / self.total_files) * 100
            lines.append(f"  Convertis          : {self.total_success} ({success_pct:.0f}%)")
        else:
            lines.append(f"  Convertis          : 0")
        lines.append(f"  Ignorés            : {self.total_skipped}")
        lines.append(f"    - Déjà existants : {len(self.skipped_existing)}")
        lines.append(f"    - Mot de passe   : {len(self.password_protected)}")
        lines.append(f"  Échecs             : {self.total_failed}")
        lines.append("")

        # Statistiques de taille
        if self.total_source_bytes > 0:
            lines.append("VOLUMES")
            lines.append(sep_light)
            lines.append(f"  Taille sources     : {self._format_size(self.total_source_bytes)}")
            lines.append(f"  Taille PDF générés : {self._format_size(self.total_dest_bytes)}")
            if self.total_dest_bytes > 0:
                ratio = (self.total_dest_bytes / self.total_source_bytes) * 100
                lines.append(f"  Ratio              : {ratio:.0f}%")
            lines.append(f"  Temps conversion   : {self._format_duration(self.total_duration)}")
            if self.total_success > 0:
                avg_time = self.total_duration / self.total_success
                lines.append(f"  Temps moyen/fichier: {avg_time:.2f}s")
            lines.append("")

        # Détail par type de fichier
        if self.stats_by_type:
            lines.append("DÉTAIL PAR TYPE")
            lines.append(sep_light)

            # Trier par nombre de fichiers décroissant
            sorted_types = sorted(
                self.stats_by_type.items(),
                key=lambda x: x[1].count,
                reverse=True
            )

            for ext, stats in sorted_types:
                status_parts = []
                if stats.success > 0:
                    status_parts.append(f"{stats.success} OK")
                if stats.failed > 0:
                    status_parts.append(f"{stats.failed} échec")
                if stats.skipped_exists > 0:
                    status_parts.append(f"{stats.skipped_exists} existant")
                if stats.skipped_password > 0:
                    status_parts.append(f"{stats.skipped_password} mdp")

                status_str = ", ".join(status_parts) if status_parts else "aucun"
                size_str = self._format_size(stats.source_size_bytes)

                lines.append(f"  {ext:8} : {stats.count:4} fichiers ({size_str:>10}) -> {status_str}")

            lines.append("")

        # Détail des conversions réussies
        if self.conversions:
            lines.append("CONVERSIONS RÉUSSIES")
            lines.append(sep_light)
            for source, dest, method, duration in self.conversions:
                # Afficher le chemin relatif si possible
                try:
                    if self.source_directory and source.is_relative_to(self.source_directory):
                        rel_source = source.relative_to(self.source_directory)
                    else:
                        rel_source = source.name
                except ValueError:
                    rel_source = source.name

                # Nom du fichier destination
                dest_name = dest.name if dest else "?"

                lines.append(f"  {rel_source}")
                lines.append(f"      -> {dest_name} ({method}, {duration:.1f}s)")
            lines.append("")

        # Détail des erreurs
        if self.errors:
            lines.append("ÉCHECS DÉTAILLÉS")
            lines.append(sep_light)
            for i, (path, reason, details) in enumerate(self.errors, 1):
                lines.append(f"  [{i}] {path.name}")
                lines.append(f"      Chemin  : {path}")
                lines.append(f"      Raison  : {reason}")
                if details:
                    # Afficher la traceback complète avec indentation
                    lines.append("      Détails :")
                    for detail_line in details.strip().split("\n"):
                        lines.append(f"        {detail_line}")
                lines.append("")

        # Fichiers protégés par mot de passe
        if self.password_protected:
            lines.append("FICHIERS PROTÉGÉS PAR MOT DE PASSE")
            lines.append(sep_light)
            for path in self.password_protected:
                lines.append(f"  - {path}")
            lines.append("")

        # Pied de page
        lines.append(sep)
        lines.append(f"Rapport généré le {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(sep)

        return "\n".join(lines)

    def save(self, output_dir: Path | None = None) -> Path | None:
        """
        Sauvegarde le rapport dans un fichier.

        Args:
            output_dir: Répertoire de sortie (par défaut: répertoire source)

        Returns:
            Chemin du fichier créé ou None si erreur
        """
        if output_dir is None:
            output_dir = self.output_directory or self.source_directory or Path.cwd()

        try:
            output_dir.mkdir(parents=True, exist_ok=True)
            timestamp = self.start_time.strftime("%Y%m%d_%H%M%S")
            report_path = output_dir / f"conversion_report_{timestamp}.txt"

            content = self.generate()
            with open(report_path, "w", encoding="utf-8") as f:
                f.write(content)

            return report_path

        except Exception:
            return None

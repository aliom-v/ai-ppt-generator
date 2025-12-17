"""导出模块

支持将 PPT 导出为其他格式（PDF、图片等）。
"""
import os
import subprocess
import tempfile
import shutil
from abc import ABC, abstractmethod
from pathlib import Path
from typing import List, Optional, Dict, Any

from utils.logger import get_logger

logger = get_logger("export")


class ExportError(Exception):
    """导出错误"""
    pass


class Exporter(ABC):
    """导出器基类"""

    @abstractmethod
    def export(self, input_path: str, output_path: str) -> str:
        """执行导出"""
        pass

    @abstractmethod
    def is_available(self) -> bool:
        """检查导出器是否可用"""
        pass


class LibreOfficeExporter(Exporter):
    """LibreOffice 导出器

    使用 LibreOffice 将 PPTX 转换为 PDF。
    需要安装 LibreOffice。
    """

    def __init__(self, libreoffice_path: str = None):
        self._path = libreoffice_path or self._find_libreoffice()

    def _find_libreoffice(self) -> Optional[str]:
        """查找 LibreOffice 可执行文件"""
        # 常见路径
        candidates = [
            "libreoffice",
            "soffice",
            "/usr/bin/libreoffice",
            "/usr/bin/soffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]

        for path in candidates:
            if shutil.which(path):
                return path

        return None

    def is_available(self) -> bool:
        return self._path is not None

    def export(self, input_path: str, output_path: str) -> str:
        """导出为 PDF"""
        if not self.is_available():
            raise ExportError("LibreOffice 未安装或不可用")

        input_path = os.path.abspath(input_path)
        output_dir = os.path.dirname(os.path.abspath(output_path))

        # 确保输出目录存在
        Path(output_dir).mkdir(parents=True, exist_ok=True)

        try:
            # 使用 LibreOffice 转换
            cmd = [
                self._path,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", output_dir,
                input_path,
            ]

            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120,
            )

            if result.returncode != 0:
                raise ExportError(f"LibreOffice 转换失败: {result.stderr}")

            # LibreOffice 会使用原文件名
            expected_output = os.path.join(
                output_dir,
                Path(input_path).stem + ".pdf"
            )

            # 如果输出路径不同，重命名
            if expected_output != output_path and os.path.exists(expected_output):
                shutil.move(expected_output, output_path)

            if not os.path.exists(output_path):
                raise ExportError("导出文件未生成")

            logger.info(f"PDF 导出成功: {output_path}")
            return output_path

        except subprocess.TimeoutExpired:
            raise ExportError("导出超时")
        except Exception as e:
            raise ExportError(f"导出失败: {e}")


class ImageExporter(Exporter):
    """图片导出器

    将 PPT 每页导出为图片。
    需要安装 pdf2image 和 poppler。
    """

    def __init__(self, dpi: int = 150, format: str = "png"):
        self.dpi = dpi
        self.format = format
        self._pdf_exporter = LibreOfficeExporter()

    def is_available(self) -> bool:
        try:
            from pdf2image import convert_from_path
            return self._pdf_exporter.is_available()
        except ImportError:
            return False

    def export(self, input_path: str, output_path: str) -> str:
        """导出为图片

        output_path 应该是一个目录，图片会按页码命名。
        """
        if not self.is_available():
            raise ExportError("图片导出器不可用（需要 pdf2image 和 LibreOffice）")

        from pdf2image import convert_from_path

        # 确保输出目录存在
        output_dir = Path(output_path)
        output_dir.mkdir(parents=True, exist_ok=True)

        # 先转换为 PDF
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp_pdf = tmp.name

        try:
            self._pdf_exporter.export(input_path, tmp_pdf)

            # 转换 PDF 为图片
            images = convert_from_path(tmp_pdf, dpi=self.dpi)

            output_files = []
            for i, image in enumerate(images, 1):
                output_file = output_dir / f"slide_{i:03d}.{self.format}"
                image.save(str(output_file), self.format.upper())
                output_files.append(str(output_file))

            logger.info(f"图片导出成功: {len(output_files)} 张 -> {output_path}")
            return str(output_dir)

        finally:
            if os.path.exists(tmp_pdf):
                os.unlink(tmp_pdf)


class ThumbnailExporter:
    """缩略图导出器

    生成 PPT 首页缩略图。
    """

    def __init__(self, width: int = 400, height: int = 300):
        self.width = width
        self.height = height
        self._image_exporter = ImageExporter(dpi=72)

    def is_available(self) -> bool:
        try:
            from PIL import Image
            return self._image_exporter.is_available()
        except ImportError:
            return False

    def export(self, input_path: str, output_path: str) -> str:
        """生成缩略图"""
        if not self.is_available():
            raise ExportError("缩略图导出器不可用")

        from PIL import Image

        with tempfile.TemporaryDirectory() as tmp_dir:
            # 导出第一页图片
            self._image_exporter.export(input_path, tmp_dir)

            # 找到第一张图片
            first_image = os.path.join(tmp_dir, "slide_001.png")
            if not os.path.exists(first_image):
                raise ExportError("无法生成缩略图")

            # 调整大小
            with Image.open(first_image) as img:
                img.thumbnail((self.width, self.height), Image.Resampling.LANCZOS)
                img.save(output_path)

            logger.info(f"缩略图导出成功: {output_path}")
            return output_path


class ExportManager:
    """导出管理器

    统一管理各种导出格式。

    用法:
        manager = ExportManager()

        # 检查可用的导出格式
        print(manager.available_formats)

        # 导出为 PDF
        manager.export("input.pptx", "output.pdf", format="pdf")

        # 导出为图片
        manager.export("input.pptx", "./slides/", format="images")
    """

    def __init__(self):
        self._exporters = {
            "pdf": LibreOfficeExporter(),
            "images": ImageExporter(),
            "thumbnail": ThumbnailExporter(),
        }

    @property
    def available_formats(self) -> List[str]:
        """获取可用的导出格式"""
        return [
            fmt for fmt, exporter in self._exporters.items()
            if exporter.is_available()
        ]

    def is_format_available(self, format: str) -> bool:
        """检查格式是否可用"""
        exporter = self._exporters.get(format)
        return exporter is not None and exporter.is_available()

    def export(self, input_path: str, output_path: str, format: str = "pdf") -> str:
        """导出文件

        Args:
            input_path: 输入 PPTX 文件路径
            output_path: 输出路径
            format: 导出格式 (pdf, images, thumbnail)

        Returns:
            输出文件/目录路径
        """
        exporter = self._exporters.get(format)
        if not exporter:
            raise ExportError(f"不支持的导出格式: {format}")

        if not exporter.is_available():
            raise ExportError(f"导出格式 {format} 当前不可用")

        if not os.path.exists(input_path):
            raise ExportError(f"输入文件不存在: {input_path}")

        return exporter.export(input_path, output_path)

    def get_status(self) -> Dict[str, Any]:
        """获取导出器状态"""
        return {
            "available_formats": self.available_formats,
            "exporters": {
                fmt: {
                    "available": exporter.is_available(),
                }
                for fmt, exporter in self._exporters.items()
            }
        }


# 全局导出管理器
_export_manager: Optional[ExportManager] = None


def get_export_manager() -> ExportManager:
    """获取全局导出管理器"""
    global _export_manager
    if _export_manager is None:
        _export_manager = ExportManager()
    return _export_manager


def export_to_pdf(input_path: str, output_path: str) -> str:
    """导出为 PDF（便捷函数）"""
    return get_export_manager().export(input_path, output_path, "pdf")


def export_to_images(input_path: str, output_dir: str) -> str:
    """导出为图片（便捷函数）"""
    return get_export_manager().export(input_path, output_dir, "images")


def generate_thumbnail(input_path: str, output_path: str) -> str:
    """生成缩略图（便捷函数）"""
    return get_export_manager().export(input_path, output_path, "thumbnail")

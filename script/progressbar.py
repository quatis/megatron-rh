import time
import os
from datetime import datetime, timedelta


class ProgressBar:
    def __init__(self, total, description="Progresso", width=50):
        self.total = total
        self.current = 0
        self.description = description
        self.width = width
        self.start_time = time.time()
        self.last_update = 0
        self.filled_char = "█"
        self.empty_char = "░"
        self.partial_chars = ["▏", "▎", "▍", "▌", "▋", "▊", "▉", "█"]

        self.colors = {
            'reset': '\033[0m',
            'bold': '\033[1m',
            'dim': '\033[2m',
            'green': '\033[92m',
            'blue': '\033[94m',
            'cyan': '\033[96m',
            'yellow': '\033[93m',
            'white': '\033[97m',
            'red': '\033[91m'
        }

    def update(self, current=None, position_info=""):
        if current is not None:
            self.current = current
        else:
            self.current += 1

        # Limitar updates muito frequentes
        now = time.time()
        if now - self.last_update < 0.1 and self.current < self.total:
            return
        self.last_update = now

        self._render(position_info)

    def _render(self, position_info=""):
        if self.total == 0:
            percentage = 100
        else:
            percentage = (self.current / self.total) * 100

        progress = self.current / self.total if self.total > 0 else 1
        filled_width = progress * self.width
        filled_blocks = int(filled_width)

        bar_filled = self.filled_char * filled_blocks

        if filled_blocks < self.width and self.current < self.total:
            partial_progress = filled_width - filled_blocks
            partial_index = int(partial_progress * len(self.partial_chars))
            if partial_index < len(self.partial_chars):
                bar_filled += self.partial_chars[partial_index]
                filled_blocks += 1

        bar_empty = self.empty_char * (self.width - filled_blocks)
        bar = bar_filled + bar_empty

        elapsed_time = time.time() - self.start_time

        if self.current > 0 and self.current < self.total:
            rate = self.current / elapsed_time
            eta_seconds = (self.total - self.current) / rate
            eta = str(timedelta(seconds=int(eta_seconds)))
            speed = f"{rate:.1f}/s"
        else:
            eta = "--:--:--"
            speed = "--/s"

        elapsed = str(timedelta(seconds=int(elapsed_time)))

        if percentage == 100:
            status_color = self.colors['green']
            status_symbol = "✅"
        elif percentage >= 75:
            status_color = self.colors['blue']
            status_symbol = "🔵"
        elif percentage >= 50:
            status_color = self.colors['yellow']
            status_symbol = "🟡"
        elif percentage >= 25:
            status_color = self.colors['cyan']
            status_symbol = "🔵"
        else:
            status_color = self.colors['white']
            status_symbol = "⚪"

        print("\r" + " " * 120, end="")

        progress_line = (
            f"\r{status_symbol} "
            f"{self.colors['white']}{self.colors['bold']}{self.description}:{self.colors['reset']} "
            f"{status_color}{self.current:>3}/{self.total:<3}{self.colors['reset']} "
            f"{self.colors['cyan']}[{bar}]{self.colors['reset']} "
            f"{status_color}{percentage:5.1f}%{self.colors['reset']} "
            f"{self.colors['dim']}| {elapsed} < {eta} | {speed}{self.colors['reset']}"
        )

        if position_info:
            progress_line += f" {self.colors['yellow']}{position_info}{self.colors['reset']}"

        print(progress_line, end="", flush=True)

        if self.current >= self.total:
            print()  # Nova linha

    def finish(self, message="Concluído"):
        self.current = self.total
        self._render()
        total_time = time.time() - self.start_time
        elapsed = str(timedelta(seconds=int(total_time)))
        print(f"\n{self.colors['green']}🎉 {message}! Tempo total: {elapsed}{self.colors['reset']}")
import re
import datetime
from pathlib import Path
from typing import List, Tuple, Optional

try:
    from fit_tool.fit_file_builder import FitFileBuilder
    from fit_tool.profile.messages.file_id_message import FileIdMessage
    from fit_tool.profile.messages.workout_message import WorkoutMessage
    from fit_tool.profile.messages.workout_step_message import WorkoutStepMessage
    from fit_tool.profile.profile_type import (
        Sport, Intensity, WorkoutStepDuration,
        WorkoutStepTarget, Manufacturer, FileType
    )
except ImportError:
    print("Error: fit-tool library not found!")
    print("Install it with: pip install fit-tool")
    exit(1)


class ErgParser:
    """Parser for ERG workout files"""

    def __init__(self, filepath: str):
        self.filepath = Path(filepath)
        self.description = ""
        self.filename = ""
        self.ftp = None
        self.intervals: List[Tuple[float, float]] = []  # (minutes, watts)
        self.text_cues: List[Tuple[float, str, int]] = []  # (seconds, message, duration)

    def parse(self) -> bool:
        """Parse the ERG file"""
        try:
            with open(self.filepath, 'r') as f:
                content = f.read()

            # Parse header section
            header_match = re.search(
                r'\[COURSE HEADER\](.*?)\[END COURSE HEADER\]',
                content,
                re.DOTALL | re.IGNORECASE
            )

            if header_match:
                header = header_match.group(1)

                # Extract description
                desc_match = re.search(r'DESCRIPTION\s*=\s*(.+)', header, re.IGNORECASE)
                if desc_match:
                    self.description = desc_match.group(1).strip()

                # Extract filename
                name_match = re.search(r'FILE NAME\s*=\s*(.+)', header, re.IGNORECASE)
                if name_match:
                    self.filename = name_match.group(1).strip()

                # Extract FTP if present
                ftp_match = re.search(r'FTP\s*=\s*(\d+)', header, re.IGNORECASE)
                if ftp_match:
                    self.ftp = int(ftp_match.group(1))

            # Parse course data section
            data_match = re.search(
                r'\[COURSE DATA\](.*?)\[END COURSE DATA\]',
                content,
                re.DOTALL | re.IGNORECASE
            )

            if data_match:
                data = data_match.group(1)
                for line in data.strip().split('\n'):
                    line = line.strip()
                    if line and not line.startswith(';'):
                        parts = re.split(r'\s+', line)
                        if len(parts) >= 2:
                            minutes = float(parts[0])
                            watts = float(parts[1])
                            self.intervals.append((minutes, watts))

            # Parse text cues section (optional)
            text_match = re.search(
                r'\[COURSE TEXT\](.*?)\[END COURSE TEXT\]',
                content,
                re.DOTALL | re.IGNORECASE
            )

            if text_match:
                text = text_match.group(1)
                for line in text.strip().split('\n'):
                    line = line.strip()
                    if line and not line.startswith(';'):
                        parts = line.split('\t')
                        if len(parts) >= 2:
                            seconds = int(parts[0])
                            message = parts[1]
                            duration = int(parts[2]) if len(parts) > 2 else 5
                            self.text_cues.append((seconds, message, duration))

            return len(self.intervals) > 0

        except Exception as e:
            print(f"Error parsing ERG file: {e}")
            return False


class FitConverter:
    """Convert parsed ERG data to FIT workout file"""

    def __init__(self, erg_parser: ErgParser, ftp: int = 200):
        self.erg = erg_parser
        self.ftp = ftp if erg_parser.ftp is None else erg_parser.ftp

    def convert_to_fit(self, output_path: str) -> bool:
        """Convert ERG data to FIT workout file"""
        try:
            # Create file ID message
            file_id = FileIdMessage()
            file_id.type = FileType.WORKOUT
            file_id.manufacturer = Manufacturer.DEVELOPMENT.value
            file_id.product = 0
            file_id.time_created = round(datetime.datetime.now().timestamp() * 1000)
            file_id.serial_number = 0x12345678

            # Create workout message
            workout = WorkoutMessage()
            workout.workout_name = self.erg.filename or Path(self.erg.filepath).stem
            workout.sport = Sport.CYCLING

            # Create workout steps by processing intervals between waypoints
            steps = []

            # Process intervals between consecutive waypoints
            for i in range(len(self.erg.intervals) - 1):
                start_time, start_watts = self.erg.intervals[i]
                end_time, end_watts = self.erg.intervals[i + 1]

                # Calculate duration
                duration_minutes = end_time - start_time
                duration_seconds = duration_minutes * 60

                # Skip zero-duration intervals (duplicate time = step change marker)
                if duration_seconds < 0.1:  # Less than ~6 seconds
                    continue

                step = WorkoutStepMessage()
                step.duration_type = WorkoutStepDuration.TIME
                step.duration_time = duration_seconds * 1000  # Convert to milliseconds
                step.target_type = WorkoutStepTarget.POWER

                # Check if this is a ramp or steady interval
                # Convert watts to percentage of FTP (FIT format expects percentages)
                power_diff = abs(end_watts - start_watts)

                if power_diff > 1 and duration_seconds > 1:  # Ramp: break into steps
                    # Adaptive step size based on interval duration
                    if duration_seconds <= 30:
                        step_size = 1  # 1 second for short sprints
                    elif duration_seconds <= 120:
                        step_size = 2  # 2 seconds for medium intervals
                    elif duration_seconds <= 300:
                        step_size = 5  # 5 seconds for longer intervals
                    else:
                        step_size = 10  # 10 seconds for very long intervals

                    num_steps = max(int(duration_seconds / step_size), 1)
                    step_duration_ms = (duration_seconds * 1000) / num_steps

                    for j in range(num_steps):
                        # Calculate power for this micro-step (linear interpolation)
                        progress = (j + 1) / num_steps
                        current_watts = start_watts + (end_watts - start_watts) * progress
                        prev_watts = start_watts + (end_watts - start_watts) * (j / num_steps)

                        micro_step = WorkoutStepMessage()
                        micro_step.duration_type = WorkoutStepDuration.TIME
                        micro_step.duration_time = step_duration_ms
                        micro_step.target_type = WorkoutStepTarget.POWER

                        # Use narrow range for each micro-step to approximate ramp
                        micro_step.custom_target_power_low = int(round(prev_watts / self.ftp * 100))
                        micro_step.custom_target_power_high = int(round(current_watts / self.ftp * 100))
                        micro_step.intensity = self._determine_intensity(current_watts / self.ftp)

                        steps.append(micro_step)

                else:  # Steady: power stays the same
                    step.custom_target_power_low = int(round(start_watts / self.ftp * 100))
                    step.custom_target_power_high = int(round(start_watts / self.ftp * 100))
                    step.intensity = self._determine_intensity(start_watts / self.ftp)

                    # Add text cues that fall within this step
                    time_start = start_time * 60
                    time_end = end_time * 60
                    cues_in_step = [
                        cue[1] for cue in self.erg.text_cues
                        if time_start <= cue[0] < time_end
                    ]
                    if cues_in_step:
                        step.notes = "; ".join(cues_in_step)

                    steps.append(step)

            workout.num_valid_steps = len(steps)

            # Build FIT file
            builder = FitFileBuilder(auto_define=True, min_string_size=50)
            builder.add(file_id)
            builder.add(workout)
            builder.add_all(steps)

            fit_file = builder.build()
            fit_file.to_file(output_path)

            return True

        except Exception as e:
            print(f"Error converting to FIT: {e}")
            return False

    @staticmethod
    def _determine_intensity(power_ratio: float) -> Intensity:
        """Determine workout intensity based on power ratio to FTP"""
        if power_ratio < 0.55:
            return Intensity.WARMUP
        elif power_ratio < 0.75:
            return Intensity.ACTIVE
        elif power_ratio < 0.95:
            return Intensity.REST
        else:
            return Intensity.ACTIVE


def convert_erg_to_fit(erg_path: str, fit_path: Optional[str] = None, ftp: int = 200) -> bool:
    """
    Convert an ERG file to FIT format

    Args:
        erg_path: Path to input .erg file
        fit_path: Path to output .fit file (optional, auto-generated if None)
        ftp: Functional Threshold Power in watts (default: 200)

    Returns:
        True if conversion successful, False otherwise
    """
    # Parse ERG file
    parser = ErgParser(erg_path)
    if not parser.parse():
        print(f"Failed to parse ERG file: {erg_path}")
        return False

    # Generate output path if not provided
    if fit_path is None:
        erg_file = Path(erg_path)
        fit_path = str(erg_file.parent / f"{erg_file.stem}.fit")

    # Convert to FIT
    converter = FitConverter(parser, ftp)
    if converter.convert_to_fit(fit_path):
        print(f"Successfully converted: {erg_path} -> {fit_path}")
        return True
    else:
        print(f"Failed to convert: {erg_path}")
        return False


def batch_convert(directory: str, ftp: int = 200):
    """
    Convert all ERG files in a directory to FIT format

    Args:
        directory: Directory containing .erg files
        ftp: Functional Threshold Power in watts
    """
    dir_path = Path(directory)
    erg_files = list(dir_path.glob("*.erg"))

    if not erg_files:
        print(f"No .erg files found in {directory}")
        return

    print(f"Found {len(erg_files)} ERG files")
    success_count = 0

    for erg_file in erg_files:
        if convert_erg_to_fit(str(erg_file), ftp=ftp):
            success_count += 1

    print(f"\nConverted {success_count}/{len(erg_files)} files successfully")


if __name__ == "__main__":
    import sys

    convert_erg_to_fit(r"dev\Vejer V3.erg", r"dev\output2.fit", 360)
    sys.exit(1)
    if len(sys.argv) < 2:
        print("Usage:")
        print("  Single file: python erg_to_fit_converter.py <input.erg> [output.fit] [ftp]")
        print("  Batch mode:  python erg_to_fit_converter.py --batch <directory> [ftp]")
        print("\nExample:")
        print("  python erg_to_fit_converter.py workout.erg")
        print("  python erg_to_fit_converter.py workout.erg output.fit 250")
        print("  python erg_to_fit_converter.py --batch ./workouts 250")
        sys.exit(1)

    if sys.argv[1] == "--batch":
        directory = sys.argv[2] if len(sys.argv) > 2 else "."
        ftp = int(sys.argv[3]) if len(sys.argv) > 3 else 200
        batch_convert(directory, ftp)
    else:
        erg_path = sys.argv[1]
        fit_path = sys.argv[2] if len(sys.argv) > 2 else None
        ftp = int(sys.argv[3]) if len(sys.argv) > 3 else 200
        convert_erg_to_fit(erg_path, fit_path, ftp)

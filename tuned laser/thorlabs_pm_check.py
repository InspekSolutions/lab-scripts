import argparse
import time

from pm_control import create_powermeter


def main():
    parser = argparse.ArgumentParser(description="Standalone Thorlabs powermeter check.")
    parser.add_argument(
        "--resource",
        default="",
        help="VISA resource string (example: USB0::0x1313::0x8078::P0000001::INSTR)",
    )
    parser.add_argument("--samples", type=int, default=10, help="Number of measurements to read")
    parser.add_argument("--delay", type=float, default=0.3, help="Delay between measurements (s)")
    args = parser.parse_args()

    pm = create_powermeter(pm_type="thorlabs", thorlabs_resource=args.resource)
    try:
        pm.set_auto_range(True)
        print("Connected. Reading power...")
        for i in range(args.samples):
            power_w = pm.read_power()
            print(f"{i + 1:02d}: {power_w:.6e} W")
            time.sleep(args.delay)
    finally:
        pm.close()


if __name__ == "__main__":
    main()

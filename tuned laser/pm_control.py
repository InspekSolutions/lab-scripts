import os
from abc import ABC, abstractmethod

from Powermeter import PW_1936r


class BasePowerMeter(ABC):
    @abstractmethod
    def read_power(self) -> float:
        pass

    @abstractmethod
    def set_auto_range(self, enabled: bool) -> None:
        pass

    @abstractmethod
    def set_range(self, range_index: int) -> None:
        pass

    @abstractmethod
    def close(self) -> None:
        pass


class Newport1936RPowerMeter(BasePowerMeter):
    def __init__(self, libname: str, product_id: int = 0xCEC7):
        self.device = PW_1936r(LIBNAME=libname, product_id=product_id)
        self.set_auto_range(True)

    def read_power(self) -> float:
        return float(self.device.ask("PM:Power?").decode("utf-8"))

    def set_auto_range(self, enabled: bool) -> None:
        self.device.write("PM:AUTO 1" if enabled else "PM:AUTO 0")

    def set_range(self, range_index: int) -> None:
        self.device.write(f"PM:RAN {int(range_index)}")

    def close(self) -> None:
        self.device.close_device()


class ThorlabsVisaPowerMeter(BasePowerMeter):
    def __init__(self, resource_name: str = "", reset: bool = False):
        try:
            import pyvisa
        except Exception as exc:
            raise RuntimeError(
                "pyvisa is required for Thorlabs power meter control. "
                "Install with: pip install pyvisa pyvisa-py"
            ) from exc

        rm = pyvisa.ResourceManager()
        if resource_name:
            self.device = rm.open_resource(resource_name)
        else:
            resources = rm.list_resources()
            if not resources:
                raise RuntimeError("No VISA resources found for Thorlabs power meter.")
            self.device = rm.open_resource(resources[0])

        self.device.timeout = 5000
        if reset:
            self.device.write("*RST")

    def read_power(self) -> float:
        # Most SCPI-compatible Thorlabs PMs answer MEAS:POW?
        return float(self.device.query("MEAS:POW?").strip())

    def set_auto_range(self, enabled: bool) -> None:
        self.device.write(f"POW:RANG:AUTO {'ON' if enabled else 'OFF'}")

    def set_range(self, range_index: int) -> None:
        # Thorlabs SCPI typically uses absolute power range values rather than index.
        # Kept for interface compatibility and intentionally not implemented.
        raise NotImplementedError("Manual indexed ranges are not supported for Thorlabs VISA mode.")

    def close(self) -> None:
        self.device.close()


def create_powermeter(
    pm_type: str = "newport",
    newport_libname: str = "",
    newport_product_id: int = 0xCEC7,
    thorlabs_resource: str = "",
):
    pm = (pm_type or "newport").strip().lower()
    if pm == "newport":
        default_dll = os.path.join(os.path.dirname(os.path.abspath(__file__)), "usbdll.dll")
        libname = newport_libname or os.environ.get("NEWPORT_DLL_PATH", default_dll)
        return Newport1936RPowerMeter(libname=libname, product_id=newport_product_id)
    if pm == "thorlabs":
        resource = thorlabs_resource or os.environ.get("THORLABS_VISA_RESOURCE", "")
        return ThorlabsVisaPowerMeter(resource_name=resource)
    raise ValueError("Unsupported pm_type. Use 'newport' or 'thorlabs'.")

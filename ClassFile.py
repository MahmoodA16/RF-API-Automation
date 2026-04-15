from typing import List, Tuple, Optional


class S2P:


    def __init__(self):

        self.s11 : Optional[Tuple[float, float]] = None
        self.s21 : Optional[Tuple[float, float]] = None
        self.s12 : Optional[Tuple[float, float]] = None
        self.s22 : Optional[Tuple[float, float]] = None

        self.frequency_unit : Optional[str] = None
        self.format_type : Optional[str] = None

    def set_s11(self, magnitude:float, phase:float):
        self.s11 = (magnitude, phase)

    def set_s21(self, magnitude:float, phase:float):
        self.s21 = (magnitude, phase)

    def set_s12(self, magnitude:float, phase:float):
        self.s12 = (magnitude, phase)

    def set_s22(self, magnitude:float, phase:float):
        self.s22 = (magnitude, phase)

    
    def set_all_params(self, s11: Tuple[float,float], s21: Tuple[float,float], s12: Tuple[float,float],  s22: Tuple[float,float]):
        
        self.s11 = s11
        self.s21 = s21
        self.s12 = s12
        self.s22 = s22

    def get_s11_magnitude(self) -> Optional[float]:

        return self.s11[0] if self.s11 else None
    
    def get_s11_phase(self) -> Optional[float]:

        return self.s11[1] if self.s11 else None
    
    def get_s21_magnitude(self) -> Optional[float]:

        return self.s21[0] if self.s21 else None
    
    def get_s21_phase(self) -> Optional[float]:

        return self.s21[1] if self.s21 else None
    
    def get_s12_magnitude(self) -> Optional[float]:

        return self.s12[0] if self.s12 else None
    
    def get_s12_phase(self) -> Optional[float]:

        return self.s12[1] if self.s12 else None
    
    def get_s22_magnitude(self) -> Optional[float]:

        return self.s22[0] if self.s22 else None
    
    def get_s22_phase(self) -> Optional[float]:

        return self.s22[1] if self.s22 else None


    def __repr__(self):
        return (f"S11 = {self.s11}, S21 = {self.s21}"
                f"S12 = {self.s12}, S22 = {self.s22}")




class IdxClass:
    def __init__(self, frequency: float, psource: float, targetGamma_1f0: float, targetPhase_1f0: float, targetGamma_2f0: float = None, targetPhase_2f0: float = None, targetGamma_3f0: float = None, targetPhase_3f0: float = None):
        self.frequency: float = frequency
        self.psource: float = psource
        self.targetGamma_1f0: float = targetGamma_1f0
        self.targetPhase_1f0: float = targetPhase_1f0
        self.targetPhase_2f0: Optional[float] = targetPhase_2f0
        self.targetGamma_2f0: Optional[float] = targetGamma_2f0
        self.targetGamma_3f0: Optional[float] = targetGamma_3f0
        self.targetPhase_3f0: Optional[float] = targetPhase_3f0

    def __repr__(self):
        return (f"Frequency = {self.frequency}, Psource = {self.psource}, "
                f"TargetGamma_1f0 = {self.targetGamma_1f0}, TargetPhase_1f0 = {self.targetPhase_1f0}")

    def __hash__(self):
        # Make the object hashable so it can be used as dictionary key
        pre_hash_list = ['frequency', 'psource', 'targetGamma_1f0', 'targetPhase_1f0']
        for attr in ['targetGamma_2f0', 'targetGamma_3f0', 'targetPhase_2f0', 'targetPhase_3f0']:
            if (getattr(self, attr) is not None):
                pre_hash_list.append(attr)

        return hash(tuple(pre_hash_list))

    def __eq__(self, other):
        if not isinstance(other, IdxClass):
            return False
        
        if any((getattr(self, attr) is None and getattr(other, attr) is not None) or (getattr(self, attr) is not None and getattr(other, attr) is None) for attr in ['targetGamma_2f0', 'targetGamma_3f0', 'targetPhase_2f0', 'targetPhase_3f0']):
            return False

        
        if (self.frequency != other.frequency or
            self.psource != other.psource or
            self.targetGamma_1f0 != other.targetGamma_1f0 or
            self.targetPhase_1f0 != other.targetPhase_1f0 or
            self.targetGamma_2f0 != other.targetGamma_2f0 or
            self.targetPhase_2f0 != other.targetPhase_2f0 or
            self.targetGamma_3f0 != other.targetGamma_3f0 or
            self.targetPhase_3f0 != other.targetPhase_3f0
            ):
            return False
        
        
        return True
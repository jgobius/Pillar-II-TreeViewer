import pandas as pd

from typing import Union

class Company:

    def __init__(self,
                 name:str,
                 ownership:Union[str, int]) -> None:

        self.name = name
        self.ownership = ownership
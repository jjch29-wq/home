import pandas as pd

class ProjectBudget:
    def __init__(self):
        self.project_name = "롯데바이오로직스 Project PROVIDENCE 중 비파괴검사"
        self.period = "2025. 02. 20 ~ 2026. 07. 31 (428일)"
        
        # 1. 수입 (매출) - 부가세 별도
        self.revenue = 377600000
        
        # 2. 직접원가
        self.labor_cost = 192070000
        self.material_cost = 486000
        self.expense = 72670784
        
        # 3. 간접비 (일반관리비 등)
        self.indirect_cost = 32755350
        
        # OT 단가 설정 (예산서 기준)
        self.ot_rates = {
            '연장': 4000,
            '야간': 5000,
            '휴일': 7500
        }

    @property
    def total_cost(self):
        """총 원가 (매출원가 + 간접비)"""
        return self.labor_cost + self.material_cost + self.expense + self.indirect_cost

    @property
    def operating_profit(self):
        """영업이익"""
        return self.revenue - self.total_cost

    @property
    def profit_margin(self):
        """영업이익률 (%)"""
        if self.revenue == 0: return 0
        return (self.operating_profit / self.revenue) * 100

    def calculate_ot_pay(self, hours, ot_type='연장'):
        """시간외 수당 계산 (예산서 단가 적용)"""
        rate = self.ot_rates.get(ot_type, 4000)
        return int(hours * rate)

    def print_summary(self):
        print(f"[{self.project_name}]")
        print(f"공사기간: {self.period}")
        print("-" * 40)
        print(f"1. 총 수입 (매출) : {self.revenue:15,} 원")
        print(f"2. 총 원가 (지출) : {self.total_cost:15,} 원")
        print(f"   - 인건비      : {self.labor_cost:15,} 원")
        print(f"   - 재료비      : {self.material_cost:15,} 원")
        print(f"   - 경비        : {self.expense:15,} 원")
        print(f"   - 간접비      : {self.indirect_cost:15,} 원")
        print("-" * 40)
        print(f"3. 영업이익      : {self.operating_profit:15,} 원")
        print(f"4. 영업이익률    : {self.profit_margin:14.2f} %")
        print("-" * 40)
        print("※ OT 단가 (예산 기준):")
        for k, v in self.ot_rates.items():
            print(f"   - {k}근무: {v:,} 원/시간")

if __name__ == "__main__":
    budget = ProjectBudget()
    budget.print_summary()

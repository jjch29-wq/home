import win32com.client as win32
import os
import time
from PIL import Image

def combine_images(paths_str, output_filename):
    paths = [p.strip() for p in paths_str.split("|") if p.strip() and os.path.exists(p.strip())]
    if not paths:
        return None
    if len(paths) == 1:
        return paths[0]
    try:
        images = [Image.open(p) for p in paths]
        min_height = min(img.height for img in images)
        resized_images = [img.resize((int(img.width * min_height / img.height), min_height)) for img in images]
        spacing = 10
        total_width = sum(img.width for img in resized_images) + spacing * (len(resized_images) - 1)
        combined = Image.new('RGB', (total_width, min_height), (255, 255, 255))
        x_offset = 0
        for img in resized_images:
            combined.paste(img, (x_offset, 0))
            x_offset += img.width + spacing
        out_dir = os.path.join(os.path.expanduser("~"), ".gemini", "scratch")
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, output_filename)
        combined.save(out_path)
        return out_path
    except Exception as e:
        print(f"이미지 병합 오류: {e}")
        return paths[0]

def generate_hwp(data, template_path, output_path):
    hwp = None
    try:
        # HWP COM 객체 생성 (화면에 띄움)
        hwp = win32.Dispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        hwp.XHwpWindows.Item(0).Visible = True

        # 한글 문서 열기
        if not hwp.Open(template_path):
            raise Exception("한글 템플릿 파일을 여는 데 실패했습니다.")

        def move_to_top():
            hwp.HAction.Run("MoveDocBegin")

        def try_put_field(label, text):
            field_str = hwp.GetFieldList()
            if not field_str:
                return False
            fields = field_str.split('\x02')
            
            # 1. 정확히 일치
            if label in fields:
                hwp.PutFieldText(label, text)
                return True
            # 2. 공백 제거 후 일치
            label_no_space = label.replace(" ", "")
            for f in fields:
                if f.replace(" ", "") == label_no_space:
                    hwp.PutFieldText(f, text)
                    return True
            # 3. 흔히 쓰는 별명(Alias) 일치
            aliases = []
            if "활동실적" in label: aliases = ["주요활동실적", "실적"]
            elif "활동계획" in label: aliases = ["주요활동계획", "계획"]
            elif "계약상대자" in label: aliases = ["업체명", "수급업체명"]
            elif "수급업체명" in label: aliases = ["업체명"]
            
            for a in aliases:
                if a in fields:
                    hwp.PutFieldText(a, text)
                    return True
            return False

        def find_text(text):
            hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
            hwp.HParameterSet.HFindReplace.FindString = text
            hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
            hwp.HParameterSet.HFindReplace.Direction = hwp.FindDir("Forward")
            return hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)

        def insert_text(text):
            hwp.HAction.Run("SelectAll")
            hwp.HAction.Run("Delete")
            hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
            hwp.HParameterSet.HInsertText.Text = text
            hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

        def fill_table_field(label, text):
            if try_put_field(label, text):
                return # 필드 이름(누름틀)으로 성공적으로 넣었으면 종료
                
            move_to_top()
            if find_text(label):
                hwp.HAction.Run("MoveRight")
                hwp.HAction.Run("SelectAll")
                hwp.HAction.Run("Copy")
                hwp.HAction.Run("Cancel")
                hwp.HAction.Run("TableRightCell")
                insert_text(text)
            else:
                print(f"[{label}] 항목을 찾을 수 없습니다.")

        def fill_risk_assessment_row(row_label, ox_key, date_key):
            o_x = data.get(ox_key, "")
            date = data.get(date_key, "")
            
            if try_put_field(ox_key, o_x) and try_put_field(date_key, date):
                return
            
            move_to_top()
            if find_text(row_label):
                hwp.HAction.Run("MoveRight")
                hwp.HAction.Run("SelectAll")
                hwp.HAction.Run("Copy")
                hwp.HAction.Run("Cancel")
                # 우측 첫번째 셀 (실시여부)
                hwp.HAction.Run("TableRightCell")
                insert_text(o_x)
                # 우측 두번째 셀 (작성 날짜)
                hwp.HAction.Run("TableRightCell")
                insert_text(date)

        # 1. 수급업체명 입력
        if not try_put_field("수급업체명", data.get("계약상대자(업체명)", "")):
            move_to_top()
            if find_text("수급업체명"):
                hwp.HAction.Run("MoveRight")
                hwp.HAction.Run("SelectAll")
                hwp.HAction.Run("Copy")
                hwp.HAction.Run("Cancel")
                hwp.HAction.Run("TableRightCell")
                insert_text(data.get("계약상대자(업체명)", ""))

        # 2. 제 출 년 월 입력
        if not try_put_field("제 출 년 월", data.get("제출년월", "")):
            move_to_top()
            if find_text("제 출 년 월"):
                hwp.HAction.Run("MoveRight")
                hwp.HAction.Run("SelectAll")
                hwp.HAction.Run("Copy")
                hwp.HAction.Run("Cancel")
                hwp.HAction.Run("TableRightCell")
                insert_text(data.get("제출년월", ""))

        # 3. 주요 필드 자동화
        mapping = {
            "계 약 명": data.get("계약명", ""),
            "계약기간": data.get("계약기간", ""),
            "계약상대자(업체명)": data.get("계약상대자(업체명)", ""),
            "현장대리인": data.get("현장대리인", ""),
            "작업의 시작시간": data.get("작업의 시작시간", ""),
            "작업 또는 작업장 간의 연락방법": data.get("작업 또는 작업장 간의 연락방법", ""),
            "재해발생 위험시의 대피방법": data.get("재해발생 위험시의 대피방법", ""),
            "사업자와 수급인 또는 수급인 상호간의 연락방법 ": data.get("사업자와 수급인 또는 수급인 상호간의 연락방법", ""),
            "작업공정의 조정 및 협의 요청사항": data.get("작업공정의 조정 및 협의 요청사항", ""),
            "주요 활동실적": data.get("주요 활동실적", ""),
            "주요 활동계획": data.get("주요 활동계획", ""),
            "관리주관부서": data.get("관리주관부서", ""),
            "장소": data.get("장소", ""),
            "중점관리항목": data.get("위험성평가_중점관리항목", ""),
            "대책1": "✔" if data.get("감소대책_위험성제거") else " ",
            "대책2": "✔" if data.get("감소대책_공학적") else " ",
            "대책3": "✔" if data.get("감소대책_관리적") else " ",
            "대책4": "✔" if data.get("감소대책_개인보호구") else " ",
            "이행사항": data.get("개선조치_이행사항", ""),
            "사고명": data.get("아차사고_사고명", ""),
            "발생일시": data.get("아차사고_발생일시", ""),
            "아차장소": data.get("아차사고_장소", ""),
            "보고자": data.get("아차사고_보고자", ""),
            "소속": data.get("아차사고_소속", ""),
            "사고내용": data.get("아차사고_사고내용", ""),
            "원인분석": data.get("아차사고_원인분석", ""),
            "개진사항": data.get("건의_개진사항", ""),
            "제안사유": data.get("건의_제안사유", "")
        }

        for label, text in mapping.items():
            if text:
                fill_table_field(label, text)

        # 4. 위험성평가 실시 여부
        fill_risk_assessment_row("최초 위험성평가", "최초위험성평가_실시여부", "최초위험성평가_작성날짜")
        fill_risk_assessment_row("정기 위험성평가", "정기위험성평가_실시여부", "정기위험성평가_작성날짜")
        fill_risk_assessment_row("수시위험성평가", "수시위험성평가_실시여부", "수시위험성평가_작성날짜")

        # 5. 사진 삽입 (조치사진)
        img_path1 = data.get("조치사진_경로", "")
        if img_path1:
            is_multiple = "|" in img_path1
            proc_path1 = combine_images(img_path1, "temp_img1.jpg")
            if proc_path1 and os.path.exists(proc_path1):
                if hwp.MoveToField("조치사진", True, True, False):
                    hwp.HAction.Run("SelectAll")
                    hwp.HAction.Run("Delete")
                    hwp.InsertPicture(proc_path1, Embedded=True, sizeoption=3 if is_multiple else 2)

        # 6. 사진 삽입 (개선후사진)
        img_path2 = data.get("개선후사진_경로", "")
        if img_path2:
            is_multiple = "|" in img_path2
            proc_path2 = combine_images(img_path2, "temp_img2.jpg")
            if proc_path2 and os.path.exists(proc_path2):
                if hwp.MoveToField("개선후사진", True, True, False):
                    hwp.HAction.Run("SelectAll")
                    hwp.HAction.Run("Delete")
                    hwp.InsertPicture(proc_path2, Embedded=True, sizeoption=3 if is_multiple else 2)
                
        # 7. 아차사고 사진 삽입 (조치전)
        img_acha1 = data.get("아차사고_조치전사진", "")
        if img_acha1:
            is_multiple = "|" in img_acha1
            proc_path3 = combine_images(img_acha1, "temp_acha1.jpg")
            if proc_path3 and os.path.exists(proc_path3):
                if hwp.MoveToField("아차조치전", True, True, False):
                    hwp.HAction.Run("SelectAll")
                    hwp.HAction.Run("Delete")
                    hwp.InsertPicture(proc_path3, Embedded=True, sizeoption=3 if is_multiple else 2)

        # 8. 아차사고 사진 삽입 (조치후)
        img_acha2 = data.get("아차사고_조치후사진", "")
        if img_acha2:
            is_multiple = "|" in img_acha2
            proc_path4 = combine_images(img_acha2, "temp_acha2.jpg")
            if proc_path4 and os.path.exists(proc_path4):
                if hwp.MoveToField("아차조치후", True, True, False):
                    hwp.HAction.Run("SelectAll")
                    hwp.HAction.Run("Delete")
                    hwp.InsertPicture(proc_path4, Embedded=True, sizeoption=3 if is_multiple else 2)

        # 문서 저장
        hwp.SaveAs(output_path)
        return output_path

    except Exception as e:
        print(f"오류 발생: {e}")
        raise e
    finally:
        if hwp:
            hwp.Quit()

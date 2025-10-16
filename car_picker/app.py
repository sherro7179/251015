from __future__ import annotations

import sys
import time
from pathlib import Path
from typing import Dict, List

APP_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = APP_DIR.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

import streamlit as st

from car_picker import config
from car_picker.quiz import QuizEngine, ensure_index, question_to_payload
from car_picker.quiz.parser import CarRecord
from car_picker.quiz.store import append_history, build_session_entry, load_history
from car_picker.quiz.utils import ensure_thumbnail


st.set_page_config(
    page_title="Car Picker Quiz",
    page_icon=":racing_car:",
    layout="wide",
)

DIFFICULTY_CHOICES = [
    ("easy", "Easy · make only"),
    ("medium", "Medium · make and model"),
    ("hard", "Hard · make, model, and year"),
]

SESSION_LENGTH_MIN = 5
SESSION_LENGTH_MAX = 40


@st.cache_resource(show_spinner=False)
def _load_cached_records() -> List[CarRecord]:
    return ensure_index(config.DATA_DIR, config.METADATA_PATH)


def load_records(rebuild: bool = False) -> List[CarRecord]:
    status_placeholder = st.empty()

    if rebuild:
        with st.spinner("Rebuilding metadata index..."):
            ensure_index(
                config.DATA_DIR,
                config.METADATA_PATH,
                rebuild=True,
                log_fn=lambda message: status_placeholder.info(message),
            )
        status_placeholder.empty()
        _load_cached_records.clear()

    with st.spinner("Loading car metadata..."):
        records = _load_cached_records()

    return records


def get_quiz_state() -> Dict[str, object]:
    if "quiz" not in st.session_state:
        st.session_state.quiz = {
            "questions": [],
            "current": 0,
            "score": 0,
            "answers": [],
            "finished": False,
            "started_at": None,
            "seed": None,
            "difficulty": config.DEFAULT_DIFFICULTY,
            "session_length": config.DEFAULT_SESSION_LENGTH,
            "last_feedback": None,
            "history_logged": False,
        }
    return st.session_state.quiz


def reset_quiz_state() -> None:
    st.session_state.quiz = {
        "questions": [],
        "current": 0,
        "score": 0,
        "answers": [],
        "finished": False,
        "started_at": None,
        "seed": None,
        "difficulty": config.DEFAULT_DIFFICULTY,
        "session_length": config.DEFAULT_SESSION_LENGTH,
        "last_feedback": None,
        "history_logged": False,
    }


def start_new_quiz(records: List[CarRecord], difficulty: str, session_length: int) -> None:
    try:
        engine = QuizEngine(records, difficulty=difficulty, session_length=session_length)
        questions = [question_to_payload(item) for item in engine.build_session()]
    except Exception as exc:  # pylint: disable=broad-except
        st.error(f"Could not prepare a new quiz session: {exc}")
        return

    reset_quiz_state()
    quiz_state = get_quiz_state()
    quiz_state.update(
        {
            "questions": questions,
            "current": 0,
            "score": 0,
            "answers": [],
            "finished": False,
            "started_at": time.time(),
            "seed": engine.seed,
            "difficulty": difficulty,
            "session_length": session_length,
            "last_feedback": None,
            "history_logged": False,
        }
    )


def record_answer(selected_option: Dict[str, object]) -> None:
    quiz_state = get_quiz_state()
    if quiz_state["finished"]:
        return

    current_index = quiz_state["current"]
    if current_index >= len(quiz_state["questions"]):
        return

    question = quiz_state["questions"][current_index]
    is_correct = bool(selected_option["is_correct"])
    quiz_state["answers"].append(
        {
            "question_id": question["id"],
            "selected_label": selected_option["label"],
            "is_correct": is_correct,
            "correct_label": question["correct_label"],
        }
    )
    if is_correct:
        quiz_state["score"] += 1

    quiz_state["last_feedback"] = {
        "question": question,
        "selected": selected_option,
        "is_correct": is_correct,
    }

    quiz_state["current"] += 1
    if quiz_state["current"] >= len(quiz_state["questions"]):
        quiz_state["finished"] = True


def maybe_log_history() -> None:
    quiz_state = get_quiz_state()
    if quiz_state["history_logged"] or not quiz_state["finished"]:
        return

    duration = 0.0
    if quiz_state["started_at"] is not None:
        duration = time.time() - quiz_state["started_at"]

    entry = build_session_entry(
        score=quiz_state["score"],
        total_questions=len(quiz_state["questions"]),
        difficulty=quiz_state["difficulty"],
        duration_seconds=duration,
        seed=quiz_state["seed"] or 0,
    )
    append_history(entry)
    quiz_state["history_logged"] = True


def render_sidebar(records_available: bool) -> Dict[str, object]:
    with st.sidebar:
        st.header("Session settings")

        difficulty_labels = {key: label for key, label in DIFFICULTY_CHOICES}
        difficulty_keys = [item[0] for item in DIFFICULTY_CHOICES]
        default_index = max(difficulty_keys.index(config.DEFAULT_DIFFICULTY), 0)

        difficulty_choice = st.selectbox(
            "Difficulty",
            options=difficulty_keys,
            format_func=lambda key: difficulty_labels[key],
            index=default_index,
        )

        session_length = st.slider(
            "Number of questions",
            min_value=SESSION_LENGTH_MIN,
            max_value=SESSION_LENGTH_MAX,
            value=config.DEFAULT_SESSION_LENGTH,
            step=5,
        )

        st.markdown("---")

        rebuild_index = st.button("Rebuild metadata index")

        start_requested = st.button(
            "Start new quiz",
            type="primary",
            disabled=not records_available,
        )

        st.markdown("---")
        st.subheader("Recent sessions")
        history = load_history()
        if not history:
            st.caption("No recorded sessions yet.")
        else:
            for entry in history[-5:][::-1]:
                st.caption(
                    f"{entry['created_at']} · {entry['difficulty']} · "
                    f"{entry['score']}/{entry['total_questions']} (seed {entry['seed']})"
                )

    return {
        "difficulty": difficulty_choice,
        "session_length": session_length,
        "rebuild_index": rebuild_index,
        "start_requested": start_requested,
    }


def render_question(question: Dict[str, object]) -> None:
    image_path = config.DATA_DIR / Path(question["image_path"])
    if not image_path.exists():
        st.error(f"Image missing: {image_path}")
        return

    try:
        thumbnail_path = ensure_thumbnail(image_path)
        st.image(str(thumbnail_path), use_column_width=True, caption=f"Question #{question['id']}")
    except Exception:  # pylint: disable=broad-except
        st.image(str(image_path), use_column_width=True, caption=f"Question #{question['id']}")

    st.markdown("#### Choose one answer")
    columns = st.columns(2, gap="medium")

    for idx, option in enumerate(question["options"]):
        target_column = columns[idx % 2]
        if target_column.button(
            option["label"],
            key=f"option_{question['id']}_{idx}",
            use_container_width=True,
        ):
            record_answer(option)
            st.rerun()


def render_summary() -> None:
    quiz_state = get_quiz_state()
    total = len(quiz_state["questions"])

    st.success(f"Quiz complete! Score: {quiz_state['score']} / {total}")

    duration = 0.0
    if quiz_state["started_at"] is not None:
        duration = time.time() - quiz_state["started_at"]

    st.markdown(
        f"- Difficulty: **{quiz_state['difficulty']}**\n"
        f"- Questions answered: **{total}**\n"
        f"- Duration: **{int(duration)}s**\n"
        f"- Seed: `{quiz_state['seed']}`"
    )

    detailed_rows = []
    for question, answer in zip(quiz_state["questions"], quiz_state["answers"], strict=False):
        detailed_rows.append(
            {
                "question": question["id"],
                "correct": question["correct_label"],
                "selected": answer["selected_label"],
                "result": "correct" if answer["is_correct"] else "wrong",
            }
        )

    st.dataframe(detailed_rows, use_container_width=True)
    maybe_log_history()

    if st.button("Play again", type="primary"):
        reset_quiz_state()
        st.rerun()


def main() -> None:
    st.title("Car Picker Quiz")
    st.write(
        "Look at each car photo and pick the matching answer. "
        "Harder modes require make, model, and year."
    )

    controls = render_sidebar(records_available=config.METADATA_PATH.exists())
    records = load_records(rebuild=controls["rebuild_index"])

    if not records:
        st.warning("No images available. Add files under `car_picker/data/` and rebuild the index.")
        st.stop()

    quiz_state = get_quiz_state()

    if controls["start_requested"]:
        start_new_quiz(
            records=records,
            difficulty=controls["difficulty"],
            session_length=controls["session_length"],
        )
        st.rerun()

    if not quiz_state["questions"]:
        st.info("Configure your session on the left and press 'Start new quiz' to begin.")
        st.stop()

    if quiz_state["finished"]:
        render_summary()
        st.stop()

    st.progress(quiz_state["current"] / len(quiz_state["questions"]))
    current_question = quiz_state["questions"][quiz_state["current"]]
    render_question(current_question)


if __name__ == "__main__":
    main()

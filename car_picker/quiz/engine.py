from __future__ import annotations

import random
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Sequence

from .. import config
from .parser import CarRecord
from .utils import stable_shuffle


@dataclass(frozen=True)
class QuestionOption:
    label: str
    make: str
    model: str
    year: int
    is_correct: bool


@dataclass(frozen=True)
class Question:
    id: int
    record: CarRecord
    options: Sequence[QuestionOption]

    @property
    def correct_label(self) -> str:
        for option in self.options:
            if option.is_correct:
                return option.label
        raise ValueError("Question has no correct option")


class QuizEngine:
    def __init__(
        self,
        records: Sequence[CarRecord],
        difficulty: str = config.DEFAULT_DIFFICULTY,
        session_length: int = config.DEFAULT_SESSION_LENGTH,
        seed: int | None = None,
    ) -> None:
        if difficulty not in config.DIFFICULTY_LEVELS:
            raise ValueError(f"Unsupported difficulty: {difficulty}")

        if session_length <= 0:
            raise ValueError("Session length must be positive")

        if len(records) < session_length:
            raise ValueError("Not enough unique records to build the session")

        self.records = list(records)
        self.difficulty = difficulty
        self.session_length = session_length
        self.seed = seed if seed is not None else random.randint(0, 1_000_000)
        self.rng = random.Random(self.seed)

        self._records_by_make: Dict[str, List[CarRecord]] = {}
        self._records_by_make_model: Dict[tuple[str, str], List[CarRecord]] = {}
        for record in self.records:
            self._records_by_make.setdefault(record.make, []).append(record)
            key = (record.make, record.model)
            self._records_by_make_model.setdefault(key, []).append(record)

    def _label_for(self, record: CarRecord) -> str:
        if self.difficulty == "easy":
            return record.label_easy
        if self.difficulty == "medium":
            return record.label_medium
        return record.label_hard

    def _build_candidate_buckets(self, record: CarRecord) -> List[List[CarRecord]]:
        same_make = [
            candidate
            for candidate in self._records_by_make.get(record.make, [])
            if candidate.key != record.key
        ]
        same_model = [
            candidate
            for candidate in self._records_by_make_model.get((record.make, record.model), [])
            if candidate.key != record.key
        ]
        similar_year = [
            candidate
            for candidate in self.records
            if candidate.key != record.key
            and candidate.make == record.make
            and abs(candidate.year - record.year) <= 2
        ]
        same_body = [
            candidate
            for candidate in self.records
            if candidate.key != record.key and candidate.body_style == record.body_style
        ]
        everything_else = [candidate for candidate in self.records if candidate.key != record.key]

        if self.difficulty == "easy":
            # We only care about different manufacturers to keep labels unique.
            different_make = [
                candidate for candidate in everything_else if candidate.make != record.make
            ]
            return [different_make, everything_else]

        if self.difficulty == "medium":
            diff_model_same_make = [
                candidate
                for candidate in same_make
                if candidate.model != record.model
            ]
            return [
                diff_model_same_make,
                same_make,
                same_body,
                everything_else,
            ]

        # hard
        same_model_diff_year = [
            candidate for candidate in same_model if candidate.year != record.year
        ]
        return [
            same_model_diff_year,
            similar_year,
            same_make,
            same_body,
            everything_else,
        ]

    def _pick_distractors(self, record: CarRecord) -> List[CarRecord]:
        required = 9
        chosen: List[CarRecord] = []
        seen_labels = {self._label_for(record)}

        buckets = self._build_candidate_buckets(record)
        for bucket in buckets:
            for candidate in stable_shuffle(bucket, self.rng):
                label = self._label_for(candidate)
                if label in seen_labels:
                    continue
                chosen.append(candidate)
                seen_labels.add(label)
                if len(chosen) == required:
                    return chosen

        # Fallback across all records if we still do not have enough.
        for candidate in stable_shuffle(self.records, self.rng):
            if len(chosen) == required:
                break
            label = self._label_for(candidate)
            if candidate.key == record.key or label in seen_labels:
                continue
            chosen.append(candidate)
            seen_labels.add(label)

        if len(chosen) < required:
            raise RuntimeError("Unable to assemble sufficient distractors")

        return chosen

    def _build_question(self, index: int, record: CarRecord) -> Question:
        correct_label = self._label_for(record)
        distractors = self._pick_distractors(record)

        options = [
            QuestionOption(
                label=correct_label,
                make=record.make,
                model=record.model,
                year=record.year,
                is_correct=True,
            )
        ]

        for distractor in distractors:
            options.append(
                QuestionOption(
                    label=self._label_for(distractor),
                    make=distractor.make,
                    model=distractor.model,
                    year=distractor.year,
                    is_correct=False,
                )
            )

        shuffled_options = stable_shuffle(options, self.rng)
        return Question(id=index, record=record, options=shuffled_options)

    def build_session(self) -> List[Question]:
        selection = stable_shuffle(self.records, self.rng)[: self.session_length]
        questions = []
        for idx, record in enumerate(selection, 1):
            questions.append(self._build_question(idx, record))
        return questions


def question_to_payload(question: Question) -> Dict[str, object]:
    return {
        "id": question.id,
        "image_path": question.record.path,
        "make": question.record.make,
        "model": question.record.model,
        "year": question.record.year,
        "body_style": question.record.body_style,
        "drivetrain": question.record.drivetrain,
        "correct_label": question.correct_label,
        "options": [
            {
                "label": option.label,
                "make": option.make,
                "model": option.model,
                "year": option.year,
                "is_correct": option.is_correct,
            }
            for option in question.options
        ],
    }

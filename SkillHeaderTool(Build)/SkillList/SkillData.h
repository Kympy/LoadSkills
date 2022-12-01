#pragma once
#include <iostream>
#include <map>
using namespace std;

const enum class SkillType {
	Wonhan, Barricade, Doctor, HeavyArmor, 
	CriticalHit, Awakening, FHFW, Gravity, 
	Beadsman, CursedDoll, Safe, MaxMana, 
	Rocket, GhostMagic, PowerShield, Adrenaline, 
	EdgeMace, BackAttack, HeadAttack, Absorb, 
	EdgeKnife, FastFast, MastersPower, FocusOnMe, 
	ManaFlow, Bombman, AngleBless, PowerFist, 
	BrokenBone, Runner, Thunder, Bully, 
	StrongMind, BreakShield, 
};

const map<SkillType, string> SkillDictionary = {
	{ SkillType::Wonhan, "원한" },
	{ SkillType::Barricade, "바리케이드" },
	{ SkillType::Doctor, "전문의" },
	{ SkillType::HeavyArmor, "중갑 착용" },
	{ SkillType::CriticalHit, "급소 타격" },
	{ SkillType::Awakening, "각성" },
	{ SkillType::FHFW, "선수필승" },
	{ SkillType::Gravity, "질량 증가" },
	{ SkillType::Beadsman, "구슬동자" },
	{ SkillType::CursedDoll, "저주받은 인형" },
	{ SkillType::Safe, "위기 모면" },
	{ SkillType::MaxMana, "최대 마나 증가" },
	{ SkillType::Rocket, "추진력" },
	{ SkillType::GhostMagic, "강령술" },
	{ SkillType::PowerShield, "강화 방패" },
	{ SkillType::Adrenaline, "아드레날린" },
	{ SkillType::EdgeMace, "예리한 둔기" },
	{ SkillType::BackAttack, "기습의 대가" },
	{ SkillType::HeadAttack, "타격의 대가" },
	{ SkillType::Absorb, "정기 흡수" },
	{ SkillType::EdgeKnife, "정밀 단도" },
	{ SkillType::FastFast, "속전속결" },
	{ SkillType::MastersPower, "달인의 저력" },
	{ SkillType::FocusOnMe, "시선 집중" },
	{ SkillType::ManaFlow, "마나의 흐름" },
	{ SkillType::Bombman, "폭발물 전문가" },
	{ SkillType::AngleBless, "여신의 가호" },
	{ SkillType::PowerFist, "분쇄의 주먹" },
	{ SkillType::BrokenBone, "부러진 뼈" },
	{ SkillType::Runner, "탈출의 명수" },
	{ SkillType::Thunder, "번개의 분노" },
	{ SkillType::Bully, "약자 무시" },
	{ SkillType::StrongMind, "굳은 의지" },
	{ SkillType::BreakShield, "쉴드 관통" },
};

const int maxCount = 34;
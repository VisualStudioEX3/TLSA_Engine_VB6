#include <windows.h>
#include <XInput.h>

#define EXPORTREAL extern "C" __declspec(dllexport) double __stdcall
#define EXPORTSTRING extern "C" __declspec(dllexport) LPSTR __stdcall

EXPORTREAL setRumble(double index, double left, double right)
{
	XINPUT_VIBRATION vibration;

	vibration.wLeftMotorSpeed = left;
	vibration.wRightMotorSpeed = right;

	XInputSetState(index,&vibration);

	return index;
}

EXPORTREAL leftTrigger(double index)
{
	XINPUT_STATE state;

	XInputGetState(index,&state);
	return state.Gamepad.bLeftTrigger;
}

EXPORTREAL rightTrigger(double index)
{
	XINPUT_STATE state;

	XInputGetState(index,&state);
	return state.Gamepad.bRightTrigger;
}

EXPORTREAL leftThumbX(double index)
{
	XINPUT_STATE state;

	XInputGetState(index,&state);
	return state.Gamepad.sThumbLX;
}

EXPORTREAL leftThumbY(double index)
{
	XINPUT_STATE state;

	XInputGetState(index,&state);
	return state.Gamepad.sThumbLY;
}

EXPORTREAL rightThumbX(double index)
{
	XINPUT_STATE state;

	XInputGetState(index,&state);
	return state.Gamepad.sThumbRX;
}

EXPORTREAL rightThumbY(double index)
{
	XINPUT_STATE state;

	XInputGetState(index,&state);
	return state.Gamepad.sThumbRY;
}

EXPORTREAL getButtonState(double index)
{
	XINPUT_STATE state;

	XInputGetState(index,&state);
	return state.Gamepad.wButtons;
}

EXPORTREAL checkButton(double index, double button)
{
	WORD buttonWord;
	XINPUT_STATE state;

	buttonWord = button;

	XInputGetState(index,&state);
	return (state.Gamepad.wButtons & buttonWord) ? 1 : 0;
}

EXPORTREAL getCtrlState(double index)
{
	XINPUT_STATE state;

	return XInputGetState(index,&state);
}

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
					 )
{
	XINPUT_VIBRATION vibration;

	vibration.wLeftMotorSpeed = 0;
	vibration.wRightMotorSpeed = 0;

	XInputSetState(0,&vibration);

	return TRUE;
}
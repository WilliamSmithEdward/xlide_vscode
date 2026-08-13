// Members a VBA UserForm has that the Microsoft Forms type library does not
// carry, so `Me.` in a form's code-behind can offer them.
//
// A form object is MSForms.UserForm plus what VBA wraps around it, and the
// difference is not small: Show and Hide - the two members form code uses most
// - are absent from the MSForms 2.0 type library entirely, as are Name, Tag and
// the position and visibility properties every form manipulates.
//
// Verified against a real form rather than assumed. A four-control UserForm was
// built in Excel through the VBE, and each candidate was asked for by name on a
// live instance via CallByName; VBA answers 438 for a member that does not
// exist and 0 (or an argument error) for one that does:
//
//   Caption=0  Name=0  Tag=0  Left=0  Top=0  Width=0  Height=0  Visible=0
//   Enabled=0  StartUpPosition=0  ShowModal=393  RightToLeft=0
//   HelpContextID=0  Hide()=0  Move()=450  Show(vbModeless)=0
//   NoSuchMemberXyz=438  NoSuchMethodXyz()=438
//
// Only what that run proved is listed here. PrintForm is documented on the
// UserForm object but was left out of the probe (it prints), and so is left out
// of this list.
//
// Completion only, like the rest of the forms metadata: the surface a form
// exposes is not exhaustively known, so absence of a name never becomes a
// diagnostic about form code.

import type { MsFormsMember } from './msformsReferenceMembers';

export const VBA_USERFORM_EXTENDER_MEMBERS: readonly MsFormsMember[] = [
	{ name: 'Show', kind: 'method', returns: 'void' },
	{ name: 'Hide', kind: 'method', returns: 'void' },
	{ name: 'Move', kind: 'method', returns: 'void' },
	{ name: 'Name', kind: 'property', returns: 'String', readOnly: true },
	{ name: 'Tag', kind: 'property', returns: 'String' },
	{ name: 'Left', kind: 'property', returns: 'Single' },
	{ name: 'Top', kind: 'property', returns: 'Single' },
	{ name: 'Width', kind: 'property', returns: 'Single' },
	{ name: 'Height', kind: 'property', returns: 'Single' },
	{ name: 'Visible', kind: 'property', returns: 'Boolean' },
	{ name: 'StartUpPosition', kind: 'property', returns: 'Long' },
	{ name: 'ShowModal', kind: 'property', returns: 'Boolean' },
	{ name: 'RightToLeft', kind: 'property', returns: 'Boolean' },
	{ name: 'HelpContextID', kind: 'property', returns: 'Long' },
];

/** The forms type whose surface VBA extends: `MSForms.UserForm`. */
export const VBA_USERFORM_TYPE = 'MSForms.UserForm';

/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_TRIBOOL_H
#define _WXAUTOEXCEL_TRIBOOL_H

#include <wx/variant.h>

namespace wxAutoExcel {
/**
    @brief Tri-state boolean.
*/
    class wxXlTribool
    {
    public:
        /*!  @brief Possible states of wxXlTribool.
        */
        enum States {
            tb_false = 0, /*!< false */
            tb_true = 1, /*!< true */
            tb_default  /*!< default / undetermined */
        };

        /**
        Creates a wxXlTribool in a default state
        */
        wxXlTribool(States state = tb_default)
            : m_state(state)
        { }

        /**
        Creates a wxXlTribool containing a bool.
        */
        wxXlTribool(bool b)
        { m_state = b ? tb_true : tb_false; }

        /**
        Returns the current state.
        */
        States GetState() const
        { return m_state; }

        /**
        Returns true if the tribool is in the default state.
        */
        bool IsDefault() const
        { return m_state == tb_default; }

        /**
        Returns true if the tribool contains true.
        */
        bool IsTrue() const
        { return m_state == tb_true; }

        /**
        Returns true if the tribool contains false.
        */
        bool IsFalse() const
        { return m_state == tb_false; }

        /**
        Assigns the bool value.
        */
        wxXlTribool& operator=(bool b)
        {
            m_state = b ? tb_true : tb_false;
            return *this;
        }

        /**
        If the variant contains a bool, assigns the bool value else sets the state to tb_default
        */
        wxXlTribool& operator=(const wxVariant& v)
        {
            if ( v.GetType() == wxS("bool") )
                *this = v.GetBool();
            else
                m_state = tb_default;
            return *this;
        }

    private:
        States m_state;
    };


    /**
    @brief XlTribool instance with default state.
    */
    extern WXDLLIMPEXP_DATA_WXAUTOEXCEL(wxXlTribool) wxDefaultXlTribool;

} // namespace wxAutoExcel

#endif // _WXAUTOEXCEL_TRIBOOL_H
